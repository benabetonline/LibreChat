const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const PowerPointToolV2 = require('./PowerPointToolV2');

const EMU = 914400;
const inch = (n) => Math.round(n * EMU);

const powerpointJsonSchema = {
  type: 'object',
  properties: {
    action: { type: 'string', enum: ['create_powerpoint'] },
    file_name: { type: 'string', description: 'Output .pptx filename.' },
    title: { type: 'string' },
    theme: {
      type: 'string',
      enum: ['auto', 'military', 'tech', 'executive', 'academic', 'minimal'],
      description: 'Visual theme. Use auto unless the user requests a style.',
    },
    slides: {
      type: 'array',
      minItems: 1,
      maxItems: 20,
      description: 'Keep each slide concise and visual.',
      items: {
        type: 'object',
        properties: {
          layout: {
            type: 'string',
            enum: ['auto', 'cover', 'cards', 'split', 'timeline', 'section', 'summary', 'content'],
          },
          title: { type: 'string' },
          subtitle: { type: 'string' },
          paragraphs: { type: 'array', maxItems: 3, items: { type: 'string' } },
          bullets: { type: 'array', maxItems: 6, items: { type: 'string' } },
        },
        required: ['title'],
      },
    },
  },
  required: ['action', 'slides'],
};

const THEMES = {
  military: {
    dark: '102019',
    dark2: '19352A',
    accent: 'C7A64A',
    accent2: '6F8F7C',
    light: 'F4F2EA',
    surface: 'E8E7DE',
    text: '17211C',
    white: 'FFFFFF',
    muted: 'BFCBC4',
  },
  tech: {
    dark: '10172B',
    dark2: '172342',
    accent: '29C6D1',
    accent2: '7257D9',
    light: 'F4F7FB',
    surface: 'E7EDF7',
    text: '17213B',
    white: 'FFFFFF',
    muted: 'BFCBE1',
  },
  executive: {
    dark: '182330',
    dark2: '263748',
    accent: 'C8924B',
    accent2: '6C879D',
    light: 'F7F6F2',
    surface: 'EBE9E3',
    text: '1E2934',
    white: 'FFFFFF',
    muted: 'C7D0D8',
  },
  academic: {
    dark: '17324D',
    dark2: '244B6B',
    accent: 'D19A3D',
    accent2: '5B8DB8',
    light: 'F7F5EF',
    surface: 'E7EDF2',
    text: '183047',
    white: 'FFFFFF',
    muted: 'C4D1DC',
  },
  minimal: {
    dark: '1D1D1F',
    dark2: '333336',
    accent: '5B67F1',
    accent2: '8C94F6',
    light: 'FAFAFA',
    surface: 'EEEEF1',
    text: '1D1D1F',
    white: 'FFFFFF',
    muted: 'C7C7CC',
  },
};

const escapeXml = (value) =>
  String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');

const safeColor = (value, fallback) => (/^[0-9A-F]{6}$/i.test(value || '') ? value : fallback);

const shapeXml = ({ id, name, x, y, w, h, fill, line, radius = false }) => {
  const fillXml = fill ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>` : '<a:noFill/>';
  const lineXml = line
    ? `<a:ln w="12700"><a:solidFill><a:srgbClr val="${line}"/></a:solidFill></a:ln>`
    : '<a:ln><a:noFill/></a:ln>';
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="${radius ? 'roundRect' : 'rect'}"><a:avLst/></a:prstGeom>${fillXml}${lineXml}</p:spPr></p:sp>`;
};

const ellipseXml = ({ id, name, x, y, w, h, fill, line }) => {
  const fillXml = fill ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>` : '<a:noFill/>';
  const lineXml = line
    ? `<a:ln w="19050"><a:solidFill><a:srgbClr val="${line}"/></a:solidFill></a:ln>`
    : '<a:ln><a:noFill/></a:ln>';
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>${fillXml}${lineXml}</p:spPr></p:sp>`;
};

const textBoxXml = ({
  id,
  name,
  x,
  y,
  w,
  h,
  text,
  size = 20,
  color = '1D1D1F',
  bold = false,
  align = 'l',
  valign = 't',
  font = 'Aptos',
}) => {
  const pt = Math.round(size * 100);
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" anchor="${valign}" lIns="0" rIns="0" tIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="${align}"/><a:r><a:rPr lang="es-EC" sz="${pt}"${bold ? ' b="1"' : ''}><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="${font}"/></a:rPr><a:t>${escapeXml(text)}</a:t></a:r><a:endParaRPr lang="es-EC" sz="${pt}"/></a:p></p:txBody></p:sp>`;
};

const baseSlide = (shapes) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;

const getThemeName = (requested, title = '') => {
  if (requested && requested !== 'auto' && THEMES[requested]) return requested;
  const t = String(title).toLowerCase();
  if (/militar|ej[eé]rcito|defensa|castrense|fuerza terrestre|seguridad nacional/.test(t)) return 'military';
  if (/inteligencia artificial|\bia\b|tecnolog|ciber|digital|software|datos/.test(t)) return 'tech';
  if (/educaci[oó]n|academ|capacitaci[oó]n|docente|curso/.test(t)) return 'academic';
  return 'executive';
};

const normalizeText = (value, max = 170) => {
  const text = String(value || '').replace(/\s+/g, ' ').trim();
  return text.length > max ? `${text.slice(0, max - 1).trim()}…` : text;
};

const chooseLayout = (slide, index, total) => {
  if (index === 0) return 'cover';
  if (index === total - 1 && total > 2) return 'summary';
  const explicit = slide.layout;
  if (explicit && !['auto', 'content', 'title'].includes(explicit)) return explicit;
  const title = String(slide.title || '').toLowerCase();
  const bullets = Array.isArray(slide.bullets) ? slide.bullets : [];
  const paragraphs = Array.isArray(slide.paragraphs) ? slide.paragraphs : [];
  if (/beneficio|riesgo|aplicacion|ventaja|desaf[ií]o|principio|pilar/.test(title) && bullets.length >= 3) return 'cards';
  if (/proceso|fase|etapa|ruta|evoluci[oó]n|cronolog/.test(title) && bullets.length >= 3) return 'timeline';
  if (paragraphs.length >= 2) return 'split';
  if (bullets.length >= 3 && bullets.length <= 6) return 'cards';
  return 'content';
};

const addHeader = (shapes, slide, theme, index) => {
  shapes.push(shapeXml({ id: 2, name: 'Background', x: 0, y: 0, w: inch(13.333), h: inch(7.5), fill: theme.light }));
  shapes.push(shapeXml({ id: 3, name: 'Accent', x: inch(0.62), y: inch(0.55), w: inch(0.12), h: inch(0.55), fill: theme.accent, radius: true }));
  shapes.push(textBoxXml({ id: 4, name: 'Title', x: inch(0.9), y: inch(0.43), w: inch(10.8), h: inch(0.8), text: normalizeText(slide.title, 85), size: 27, color: theme.text, bold: true }));
  shapes.push(textBoxXml({ id: 5, name: 'SlideNum', x: inch(11.8), y: inch(0.5), w: inch(0.7), h: inch(0.4), text: String(index + 1).padStart(2, '0'), size: 11, color: theme.accent, bold: true, align: 'r' }));
};

const renderCover = (slide, theme) => {
  const s = [];
  s.push(shapeXml({ id: 2, name: 'Background', x: 0, y: 0, w: inch(13.333), h: inch(7.5), fill: theme.dark }));
  s.push(shapeXml({ id: 3, name: 'Panel', x: inch(8.5), y: 0, w: inch(4.833), h: inch(7.5), fill: theme.dark2 }));
  s.push(shapeXml({ id: 4, name: 'AccentBar', x: inch(0.72), y: inch(1.1), w: inch(0.13), h: inch(4.9), fill: theme.accent, radius: true }));
  s.push(textBoxXml({ id: 5, name: 'Kicker', x: inch(1.15), y: inch(1.18), w: inch(6.8), h: inch(0.5), text: 'PRESENTACIÓN', size: 11, color: theme.accent, bold: true }));
  s.push(textBoxXml({ id: 6, name: 'Title', x: inch(1.15), y: inch(1.75), w: inch(6.9), h: inch(2.5), text: normalizeText(slide.title, 105), size: 32, color: theme.white, bold: true, valign: 'ctr' }));
  if (slide.subtitle) s.push(textBoxXml({ id: 7, name: 'Subtitle', x: inch(1.15), y: inch(4.55), w: inch(6.7), h: inch(1.0), text: normalizeText(slide.subtitle, 150), size: 16, color: theme.muted }));
  s.push(ellipseXml({ id: 8, name: 'Ring1', x: inch(9.15), y: inch(1.1), w: inch(2.8), h: inch(2.8), fill: null, line: theme.accent }));
  s.push(ellipseXml({ id: 9, name: 'Ring2', x: inch(10.1), y: inch(2.15), w: inch(2.2), h: inch(2.2), fill: null, line: theme.accent2 }));
  s.push(shapeXml({ id: 10, name: 'Block1', x: inch(9.0), y: inch(5.05), w: inch(1.25), h: inch(0.13), fill: theme.accent, radius: true }));
  s.push(shapeXml({ id: 11, name: 'Block2', x: inch(10.45), y: inch(5.05), w: inch(1.8), h: inch(0.13), fill: theme.accent2, radius: true }));
  return baseSlide(s);
};

const renderCards = (slide, theme, index) => {
  const s = [];
  addHeader(s, slide, theme, index);
  const items = (Array.isArray(slide.bullets) ? slide.bullets : []).slice(0, 6);
  if (!items.length && slide.subtitle) items.push(slide.subtitle);
  const count = Math.max(items.length, 1);
  const cols = count <= 4 ? 2 : 3;
  const rows = Math.ceil(count / cols);
  const left = 0.85;
  const top = 1.55;
  const gapX = 0.28;
  const gapY = 0.28;
  const usableW = 11.65;
  const usableH = 5.25;
  const cardW = (usableW - gapX * (cols - 1)) / cols;
  const cardH = (usableH - gapY * (rows - 1)) / rows;
  items.forEach((item, i) => {
    const c = i % cols;
    const r = Math.floor(i / cols);
    const x = left + c * (cardW + gapX);
    const y = top + r * (cardH + gapY);
    const id = 20 + i * 4;
    s.push(shapeXml({ id, name: `Card${i + 1}`, x: inch(x), y: inch(y), w: inch(cardW), h: inch(cardH), fill: theme.white, line: theme.surface, radius: true }));
    s.push(shapeXml({ id: id + 1, name: `CardAccent${i + 1}`, x: inch(x), y: inch(y), w: inch(0.09), h: inch(cardH), fill: i % 2 ? theme.accent2 : theme.accent, radius: true }));
    s.push(ellipseXml({ id: id + 2, name: `Num${i + 1}`, x: inch(x + 0.28), y: inch(y + 0.28), w: inch(0.48), h: inch(0.48), fill: i % 2 ? theme.accent2 : theme.accent }));
    s.push(textBoxXml({ id: id + 3, name: `CardText${i + 1}`, x: inch(x + 0.35), y: inch(y + 0.92), w: inch(cardW - 0.65), h: inch(cardH - 1.15), text: normalizeText(item, 115), size: count > 4 ? 15 : 17, color: theme.text, bold: true, valign: 'ctr' }));
    s.push(textBoxXml({ id: id + 100, name: `NumText${i + 1}`, x: inch(x + 0.28), y: inch(y + 0.31), w: inch(0.48), h: inch(0.3), text: String(i + 1), size: 11, color: theme.white, bold: true, align: 'ctr' }));
  });
  return baseSlide(s);
};

const renderSplit = (slide, theme, index) => {
  const s = [];
  addHeader(s, slide, theme, index);
  s.push(shapeXml({ id: 20, name: 'LeftPanel', x: inch(0.85), y: inch(1.55), w: inch(7.0), h: inch(5.15), fill: theme.white, line: theme.surface, radius: true }));
  s.push(shapeXml({ id: 21, name: 'RightPanel', x: inch(8.15), y: inch(1.55), w: inch(4.35), h: inch(5.15), fill: theme.dark2, radius: true }));
  const paragraphs = (Array.isArray(slide.paragraphs) ? slide.paragraphs : []).slice(0, 3);
  const bullets = (Array.isArray(slide.bullets) ? slide.bullets : []).slice(0, 4);
  let y = 1.9;
  [...paragraphs, ...bullets].slice(0, 5).forEach((item, i) => {
    s.push(ellipseXml({ id: 30 + i * 2, name: `Dot${i}`, x: inch(1.2), y: inch(y + 0.08), w: inch(0.18), h: inch(0.18), fill: i % 2 ? theme.accent2 : theme.accent }));
    s.push(textBoxXml({ id: 31 + i * 2, name: `Text${i}`, x: inch(1.55), y: inch(y), w: inch(5.75), h: inch(0.72), text: normalizeText(item, 130), size: 16, color: theme.text, bold: i === 0 }));
    y += 0.88;
  });
  const key = normalizeText(slide.subtitle || paragraphs[0] || bullets[0] || slide.title, 120);
  s.push(textBoxXml({ id: 50, name: 'KeyLabel', x: inch(8.65), y: inch(2.15), w: inch(3.2), h: inch(0.4), text: 'IDEA CLAVE', size: 10, color: theme.accent, bold: true }));
  s.push(textBoxXml({ id: 51, name: 'KeyText', x: inch(8.65), y: inch(2.8), w: inch(3.15), h: inch(2.25), text: key, size: 22, color: theme.white, bold: true, valign: 'ctr' }));
  s.push(shapeXml({ id: 52, name: 'KeyLine', x: inch(8.65), y: inch(5.55), w: inch(1.4), h: inch(0.08), fill: theme.accent, radius: true }));
  return baseSlide(s);
};

const renderTimeline = (slide, theme, index) => {
  const s = [];
  addHeader(s, slide, theme, index);
  const items = (Array.isArray(slide.bullets) ? slide.bullets : []).slice(0, 5);
  const y = 3.55;
  s.push(shapeXml({ id: 20, name: 'Timeline', x: inch(1.3), y: inch(y), w: inch(10.7), h: inch(0.06), fill: theme.accent2, radius: true }));
  const step = items.length > 1 ? 10.3 / (items.length - 1) : 0;
  items.forEach((item, i) => {
    const x = 1.42 + step * i;
    const id = 30 + i * 4;
    s.push(ellipseXml({ id, name: `Node${i}`, x: inch(x), y: inch(y - 0.22), w: inch(0.5), h: inch(0.5), fill: i % 2 ? theme.accent2 : theme.accent }));
    s.push(textBoxXml({ id: id + 1, name: `Step${i}`, x: inch(x - 0.02), y: inch(y - 0.12), w: inch(0.54), h: inch(0.28), text: String(i + 1), size: 10, color: theme.white, bold: true, align: 'ctr' }));
    const top = i % 2 === 0 ? 1.75 : 4.15;
    s.push(textBoxXml({ id: id + 2, name: `TimelineText${i}`, x: inch(Math.max(0.7, x - 1.0)), y: inch(top), w: inch(2.25), h: inch(1.15), text: normalizeText(item, 75), size: 14, color: theme.text, bold: true, align: 'ctr', valign: 'ctr' }));
  });
  return baseSlide(s);
};

const renderSection = (slide, theme, index) => {
  const s = [];
  s.push(shapeXml({ id: 2, name: 'Background', x: 0, y: 0, w: inch(13.333), h: inch(7.5), fill: theme.dark2 }));
  s.push(textBoxXml({ id: 3, name: 'Index', x: inch(0.9), y: inch(0.65), w: inch(2.2), h: inch(1.6), text: String(index + 1).padStart(2, '0'), size: 54, color: theme.accent, bold: true }));
  s.push(shapeXml({ id: 4, name: 'Line', x: inch(0.95), y: inch(2.35), w: inch(1.5), h: inch(0.1), fill: theme.accent, radius: true }));
  s.push(textBoxXml({ id: 5, name: 'Title', x: inch(3.2), y: inch(1.65), w: inch(8.8), h: inch(2.4), text: normalizeText(slide.title, 95), size: 34, color: theme.white, bold: true, valign: 'ctr' }));
  if (slide.subtitle) s.push(textBoxXml({ id: 6, name: 'Subtitle', x: inch(3.2), y: inch(4.35), w: inch(7.8), h: inch(0.9), text: normalizeText(slide.subtitle, 125), size: 16, color: theme.muted }));
  return baseSlide(s);
};

const renderContent = (slide, theme, index) => {
  const s = [];
  addHeader(s, slide, theme, index);
  const items = [
    ...(Array.isArray(slide.bullets) ? slide.bullets : []),
    ...(Array.isArray(slide.paragraphs) ? slide.paragraphs : []),
  ].slice(0, 6);
  let y = 1.55;
  items.forEach((item, i) => {
    const id = 20 + i * 4;
    s.push(shapeXml({ id, name: `Row${i}`, x: inch(0.9), y: inch(y), w: inch(11.6), h: inch(0.78), fill: i % 2 ? theme.surface : theme.white, radius: true }));
    s.push(textBoxXml({ id: id + 1, name: `RowNum${i}`, x: inch(1.15), y: inch(y + 0.18), w: inch(0.55), h: inch(0.35), text: String(i + 1).padStart(2, '0'), size: 11, color: theme.accent, bold: true }));
    s.push(textBoxXml({ id: id + 2, name: `RowText${i}`, x: inch(1.9), y: inch(y + 0.12), w: inch(9.9), h: inch(0.52), text: normalizeText(item, 150), size: 15, color: theme.text, bold: i === 0, valign: 'ctr' }));
    y += 0.9;
  });
  if (slide.subtitle) s.push(textBoxXml({ id: 80, name: 'Subtitle', x: inch(0.95), y: inch(6.65), w: inch(10.8), h: inch(0.35), text: normalizeText(slide.subtitle, 145), size: 10, color: theme.accent2 }));
  return baseSlide(s);
};

const renderSummary = (slide, theme, index) => {
  const s = [];
  s.push(shapeXml({ id: 2, name: 'Background', x: 0, y: 0, w: inch(13.333), h: inch(7.5), fill: theme.dark }));
  s.push(textBoxXml({ id: 3, name: 'Kicker', x: inch(0.9), y: inch(0.7), w: inch(3), h: inch(0.4), text: 'CIERRE', size: 10, color: theme.accent, bold: true }));
  s.push(textBoxXml({ id: 4, name: 'Title', x: inch(0.9), y: inch(1.15), w: inch(10.8), h: inch(1.0), text: normalizeText(slide.title || 'Conclusiones', 85), size: 30, color: theme.white, bold: true }));
  const items = (Array.isArray(slide.bullets) ? slide.bullets : []).slice(0, 3);
  const fallback = (Array.isArray(slide.paragraphs) ? slide.paragraphs : []).slice(0, 3);
  const finalItems = items.length ? items : fallback;
  finalItems.forEach((item, i) => {
    const x = 0.9 + i * 4.05;
    const id = 20 + i * 4;
    s.push(shapeXml({ id, name: `SummaryCard${i}`, x: inch(x), y: inch(2.55), w: inch(3.65), h: inch(3.25), fill: theme.dark2, line: i % 2 ? theme.accent2 : theme.accent, radius: true }));
    s.push(textBoxXml({ id: id + 1, name: `SummaryNum${i}`, x: inch(x + 0.3), y: inch(2.85), w: inch(0.7), h: inch(0.45), text: `0${i + 1}`, size: 12, color: theme.accent, bold: true }));
    s.push(textBoxXml({ id: id + 2, name: `SummaryText${i}`, x: inch(x + 0.3), y: inch(3.55), w: inch(3.0), h: inch(1.7), text: normalizeText(item, 110), size: 18, color: theme.white, bold: true, valign: 'ctr' }));
  });
  s.push(shapeXml({ id: 90, name: 'FooterLine', x: inch(0.9), y: inch(6.55), w: inch(2.1), h: inch(0.08), fill: theme.accent, radius: true }));
  s.push(textBoxXml({ id: 91, name: 'SlideNum', x: inch(11.7), y: inch(6.45), w: inch(0.8), h: inch(0.35), text: String(index + 1).padStart(2, '0'), size: 10, color: theme.muted, align: 'r' }));
  return baseSlide(s);
};

const renderSlide = (slide, index, total, theme) => {
  const layout = chooseLayout(slide, index, total);
  if (layout === 'cover') return renderCover(slide, theme);
  if (layout === 'cards') return renderCards(slide, theme, index);
  if (layout === 'split') return renderSplit(slide, theme, index);
  if (layout === 'timeline') return renderTimeline(slide, theme, index);
  if (layout === 'section') return renderSection(slide, theme, index);
  if (layout === 'summary') return renderSummary(slide, theme, index);
  return renderContent(slide, theme, index);
};

class PowerPointToolV3 extends PowerPointToolV2 {
  constructor(fields = {}) {
    super(fields);
    this.name = 'powerpoint_tool';
    this.description = `Create polished, downloadable PowerPoint presentations. Keep slides concise. Prefer visual layouts (cover, cards, split, timeline, section, summary) instead of long bullet lists. Use theme=auto unless the user specifies a style.`;
    this.schema = powerpointJsonSchema;
  }

  static get jsonSchema() {
    return powerpointJsonSchema;
  }

  async _call(data) {
    const resultText = await super._call({ ...data, action: 'create_powerpoint' });
    let result;
    try {
      result = JSON.parse(resultText);
    } catch {
      return resultText;
    }
    if (!result?.success || !result.output_file) return resultText;

    try {
      const userDirectory = this.ensureUserDirectory();
      const candidates = fs
        .readdirSync(userDirectory)
        .filter((name) => name.endsWith(`__${result.output_file}`))
        .map((name) => ({ name, mtime: fs.statSync(path.join(userDirectory, name)).mtimeMs }))
        .sort((a, b) => b.mtime - a.mtime);
      if (!candidates.length) return resultText;

      const filePath = path.join(userDirectory, candidates[0].name);
      const zip = await JSZip.loadAsync(fs.readFileSync(filePath));
      const slides = Array.isArray(data.slides) ? data.slides : [];
      const themeName = getThemeName(data.theme, data.title || slides[0]?.title);
      const theme = THEMES[themeName] || THEMES.executive;

      for (let i = 0; i < slides.length; i++) {
        const part = zip.file(`ppt/slides/slide${i + 1}.xml`);
        if (!part) continue;
        zip.file(`ppt/slides/slide${i + 1}.xml`, renderSlide(slides[i], i, slides.length, theme));
      }

      const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
      fs.writeFileSync(filePath, output);
      result.design = 'visual-v3';
      result.theme = themeName;
      return JSON.stringify(result);
    } catch (error) {
      return JSON.stringify({ ...result, design_warning: error.message });
    }
  }
}

module.exports = PowerPointToolV3;
