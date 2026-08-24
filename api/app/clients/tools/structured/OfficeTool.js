const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const JSZip = require('jszip');
const { v4: uuidv4 } = require('uuid');
const { FileContext, FileSources } = require('librechat-data-provider');
const db = require('~/models');
const ExcelTool = require('./ExcelTool');

const officeJsonSchema = {
  type: 'object',
  properties: {
    action: {
      type: 'string',
      enum: ['inspect', 'replace_text', 'batch_edit', 'create_word', 'create_powerpoint'],
      description:
        'Office action. Use inspect, replace_text or batch_edit for Excel/CSV; create_word for DOCX; create_powerpoint for PPTX.',
    },
    file_name: {
      type: 'string',
      description:
        'For Excel actions, optional uploaded filename. For Word or PowerPoint creation, optional output filename.',
    },
    search_text: {
      type: 'string',
      description: 'Text or cell value to search for when using replace_text.',
    },
    replacement_text: {
      type: 'string',
      description: 'New text or cell value that will replace search_text.',
    },
    sheet_name: {
      type: 'string',
      description:
        'Optional worksheet name. If omitted, the Excel operation may apply to the default or all worksheets depending on the action.',
    },
    column: {
      type: 'string',
      description: 'Optional Excel column letter such as F or H.',
    },
    operations: {
      type: 'array',
      description: 'Excel operations to execute in order when action is batch_edit.',
      items: {
        type: 'object',
        properties: {
          type: {
            type: 'string',
            enum: ['replace_text', 'set_cell', 'clear_cell', 'count_value'],
          },
          sheet_name: { type: 'string' },
          column: { type: 'string' },
          search_text: { type: 'string' },
          replacement_text: { type: 'string' },
          cell: { type: 'string' },
          value: { type: 'string' },
          output_cell: { type: 'string' },
        },
        required: ['type'],
      },
    },
    title: {
      type: 'string',
      description: 'Document or presentation title.',
    },
    subtitle: {
      type: 'string',
      description: 'Optional document or presentation subtitle.',
    },
    author: {
      type: 'string',
      description: 'Optional author name for generated Office files.',
    },
    sections: {
      type: 'array',
      maxItems: 60,
      description:
        'Sections for create_word. Each section can contain a heading, paragraphs, bullets and one table.',
      items: {
        type: 'object',
        properties: {
          heading: { type: 'string' },
          level: {
            type: 'number',
            minimum: 1,
            maximum: 3,
            description: 'Heading level from 1 to 3.',
          },
          paragraphs: {
            type: 'array',
            items: { type: 'string' },
          },
          bullets: {
            type: 'array',
            items: { type: 'string' },
          },
          table: {
            type: 'object',
            properties: {
              headers: {
                type: 'array',
                items: { type: 'string' },
              },
              rows: {
                type: 'array',
                items: {
                  type: 'array',
                  items: { type: 'string' },
                },
              },
            },
          },
        },
      },
    },
    slides: {
      type: 'array',
      maxItems: 30,
      description:
        'Slides for create_powerpoint. Provide a title and concise paragraphs or bullets for each slide.',
      items: {
        type: 'object',
        properties: {
          layout: {
            type: 'string',
            enum: ['title', 'content', 'section'],
            description: 'Optional slide layout hint.',
          },
          title: { type: 'string' },
          subtitle: { type: 'string' },
          paragraphs: {
            type: 'array',
            items: { type: 'string' },
          },
          bullets: {
            type: 'array',
            maxItems: 12,
            items: { type: 'string' },
          },
        },
        required: ['title'],
      },
    },
  },
  required: ['action'],
};

const WORD_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const PPT_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

const escapeXml = (value) =>
  String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');

const sanitizeFilename = (name, extension, fallback) => {
  let safe = String(name || fallback)
    .replace(/[\\/:*?"<>|\u0000-\u001F]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/g, '');

  if (!safe) {
    safe = fallback;
  }

  const lower = safe.toLowerCase();
  if (!lower.endsWith(extension)) {
    safe += extension;
  }

  return safe.slice(0, 180);
};

const textRunXml = (text, options = {}) => {
  const bold = options.bold ? '<w:b/>' : '';
  const italic = options.italic ? '<w:i/>' : '';
  const size = options.size ? `<w:sz w:val="${options.size}"/><w:szCs w:val="${options.size}"/>` : '';
  return `<w:r><w:rPr>${bold}${italic}${size}<w:lang w:val="es-EC"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
};

const wordParagraphXml = (text, options = {}) => {
  const style = options.style ? `<w:pStyle w:val="${options.style}"/>` : '';
  const align = options.align ? `<w:jc w:val="${options.align}"/>` : '';
  const spacing = options.after !== undefined ? `<w:spacing w:after="${options.after}"/>` : '<w:spacing w:after="120"/>';
  const indent = options.indent ? '<w:ind w:left="720" w:hanging="360"/>' : '';
  return `<w:p><w:pPr>${style}${align}${spacing}${indent}</w:pPr>${textRunXml(text, options)}</w:p>`;
};

const wordTableXml = (table = {}) => {
  const headers = Array.isArray(table.headers) ? table.headers : [];
  const rows = Array.isArray(table.rows) ? table.rows : [];
  const columnCount = Math.max(headers.length, ...rows.map((row) => (Array.isArray(row) ? row.length : 0)), 1);
  const width = Math.floor(9000 / columnCount);

  const cellXml = (value, isHeader = false) =>
    `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tcMar></w:tcPr>${wordParagraphXml(value, { bold: isHeader, after: 0 })}</w:tc>`;

  const normalizedRows = [];
  if (headers.length) {
    normalizedRows.push(headers.map((value) => cellXml(value, true)));
  }
  for (const row of rows) {
    const values = Array.isArray(row) ? row : [row];
    normalizedRows.push(Array.from({ length: columnCount }, (_, index) => cellXml(values[index] ?? '')));
  }

  if (!normalizedRows.length) {
    return '';
  }

  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:color="BFBFBF"/><w:left w:val="single" w:sz="4" w:color="BFBFBF"/><w:bottom w:val="single" w:sz="4" w:color="BFBFBF"/><w:right w:val="single" w:sz="4" w:color="BFBFBF"/><w:insideH w:val="single" w:sz="4" w:color="D9D9D9"/><w:insideV w:val="single" w:sz="4" w:color="D9D9D9"/></w:tblBorders></w:tblPr>${normalizedRows.map((cells) => `<w:tr>${cells.join('')}</w:tr>`).join('')}</w:tbl>`;
};

const buildWordDocumentXml = (data) => {
  const body = [];
  body.push(wordParagraphXml(data.title || 'Documento', { style: 'Title', align: 'center', after: 180 }));
  if (data.subtitle) {
    body.push(wordParagraphXml(data.subtitle, { style: 'Subtitle', align: 'center', after: 320 }));
  }

  const sections = Array.isArray(data.sections) ? data.sections : [];
  for (const section of sections) {
    if (section.heading) {
      const level = Math.min(3, Math.max(1, Number(section.level) || 1));
      body.push(wordParagraphXml(section.heading, { style: `Heading${level}`, after: 140 }));
    }
    for (const paragraph of Array.isArray(section.paragraphs) ? section.paragraphs : []) {
      body.push(wordParagraphXml(paragraph));
    }
    for (const bullet of Array.isArray(section.bullets) ? section.bullets : []) {
      body.push(wordParagraphXml(`• ${bullet}`, { indent: true }));
    }
    if (section.table) {
      body.push(wordTableXml(section.table));
      body.push(wordParagraphXml('', { after: 120 }));
    }
  }

  if (!sections.length && data.subtitle) {
    body.push(wordParagraphXml(data.subtitle));
  }

  body.push(
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>',
  );

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body.join('')}</w:body></w:document>`;
};

const wordStylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="es-EC"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:rPr><w:i/><w:color w:val="666666"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>
</w:styles>`;

const buildCoreProps = (title, author) => {
  const now = new Date().toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${escapeXml(title || '')}</dc:title><dc:creator>${escapeXml(author || 'LibreChat')}</dc:creator><cp:lastModifiedBy>${escapeXml(author || 'LibreChat')}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;
};

const buildDocx = async (data) => {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`,
  );
  zip.folder('_rels').file(
    '.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`,
  );
  const word = zip.folder('word');
  word.file('document.xml', buildWordDocumentXml(data));
  word.file('styles.xml', wordStylesXml);
  word.folder('_rels').file(
    'document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
  );
  const props = zip.folder('docProps');
  props.file('core.xml', buildCoreProps(data.title, data.author));
  props.file(
    'app.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>LibreChat</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>1.0</AppVersion></Properties>`,
  );
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
};

const pptTextShape = ({ id, name, x, y, cx, cy, paragraphs, fontSize = 2200, bold = false }) => {
  const paraXml = paragraphs
    .map((paragraph) => {
      const isBullet = Boolean(paragraph.bullet);
      const pPr = isBullet
        ? '<a:pPr marL="457200" indent="-228600"><a:buChar char="•"/></a:pPr>'
        : '<a:pPr/>';
      return `<a:p>${pPr}<a:r><a:rPr lang="es-EC" sz="${paragraph.size || fontSize}"${paragraph.bold || bold ? ' b="1"' : ''}/><a:t>${escapeXml(paragraph.text)}</a:t></a:r><a:endParaRPr lang="es-EC" sz="${paragraph.size || fontSize}"/></a:p>`;
    })
    .join('');

  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" anchor="t"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${paraXml}</p:txBody></p:sp>`;
};

const buildSlideXml = (slide, index) => {
  const shapes = [];
  const title = slide.title || `Diapositiva ${index + 1}`;
  shapes.push(
    pptTextShape({
      id: 2,
      name: 'Title',
      x: 640080,
      y: 365760,
      cx: 10911840,
      cy: 914400,
      paragraphs: [{ text: title, bold: true, size: 3000 }],
      fontSize: 3000,
      bold: true,
    }),
  );

  const bodyParagraphs = [];
  if (slide.subtitle) {
    bodyParagraphs.push({ text: slide.subtitle, bold: slide.layout === 'title', size: slide.layout === 'title' ? 2200 : 1900 });
  }
  for (const paragraph of Array.isArray(slide.paragraphs) ? slide.paragraphs : []) {
    bodyParagraphs.push({ text: paragraph, size: 1900 });
  }
  for (const bullet of Array.isArray(slide.bullets) ? slide.bullets : []) {
    bodyParagraphs.push({ text: bullet, bullet: true, size: 1900 });
  }

  if (bodyParagraphs.length) {
    shapes.push(
      pptTextShape({
        id: 3,
        name: 'Content',
        x: 822960,
        y: 1554480,
        cx: 10515600,
        cy: 4476240,
        paragraphs: bodyParagraphs,
        fontSize: 1900,
      }),
    );
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
};

const pptThemeXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;

const buildPptx = async (data) => {
  const slides = Array.isArray(data.slides) && data.slides.length
    ? data.slides
    : [{ layout: 'title', title: data.title || 'Presentación', subtitle: data.subtitle || '' }];
  const zip = new JSZip();
  const slideOverrides = slides
    .map((_, index) => `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`)
    .join('');
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>${slideOverrides}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`,
  );
  zip.folder('_rels').file(
    '.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`,
  );

  const ppt = zip.folder('ppt');
  const slideIdXml = slides.map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 2}"/>`).join('');
  ppt.file(
    'presentation.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${slideIdXml}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr/><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr lang="es-EC"/></a:lvl1pPr></p:defaultTextStyle></p:presentation>`,
  );

  const presentationRels = [
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>',
    ...slides.map((_, index) => `<Relationship Id="rId${index + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${index + 1}.xml"/>`),
  ];
  ppt.folder('_rels').file(
    'presentation.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${presentationRels.join('')}</Relationships>`,
  );

  ppt.folder('theme').file('theme1.xml', pptThemeXml);
  ppt.folder('slideMasters').file(
    'slideMaster1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Master"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><p:sldLayoutIdLst><p:sldLayoutId id="1" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="3000" b="1"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="1900"/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="es-EC"/></a:defPPr></p:otherStyle></p:txStyles></p:sldMaster>`,
  );
  ppt.folder('slideMasters').folder('_rels').file(
    'slideMaster1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>`,
  );
  ppt.folder('slideLayouts').file(
    'slideLayout1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
  );
  ppt.folder('slideLayouts').folder('_rels').file(
    'slideLayout1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
  );

  const slideFolder = ppt.folder('slides');
  const slideRelFolder = slideFolder.folder('_rels');
  slides.forEach((slide, index) => {
    slideFolder.file(`slide${index + 1}.xml`, buildSlideXml(slide, index));
    slideRelFolder.file(
      `slide${index + 1}.xml.rels`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
    );
  });

  const props = zip.folder('docProps');
  props.file('core.xml', buildCoreProps(data.title || slides[0]?.title, data.author));
  props.file(
    'app.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>LibreChat</Application><PresentationFormat>Widescreen</PresentationFormat><Slides>${slides.length}</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><ScaleCrop>false</ScaleCrop><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>1.0</AppVersion></Properties>`,
  );

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
};

class OfficeTool extends ExcelTool {
  constructor(fields = {}) {
    super(fields);
    this.name = 'excel_tool';
    this.description = `
Office document tool for the current LibreChat user.
Use Excel actions to inspect or edit XLSX/CSV files.
Use create_word to generate a downloadable Microsoft Word .docx document.
Use create_powerpoint to generate a downloadable Microsoft PowerPoint .pptx presentation.
Generated files are private to the current user and receive short-lived signed download links.
`;
    this.schema = officeJsonSchema;
  }

  static get jsonSchema() {
    return officeJsonSchema;
  }

  ensureUserDirectory() {
    if (!this.userId) {
      throw new Error('OfficeTool: missing userId.');
    }
    const uploadsRoot = path.resolve('/app/uploads');
    const userDirectory = path.resolve(uploadsRoot, String(this.userId));
    if (!userDirectory.startsWith(`${uploadsRoot}${path.sep}`)) {
      throw new Error('OfficeTool: invalid user directory.');
    }
    fs.mkdirSync(userDirectory, { recursive: true });
    return userDirectory;
  }

  async registerGeneratedFile(filePath, filename, mime, label) {
    const fileId = uuidv4();
    const bytes = fs.statSync(filePath).size;
    await db.createFile(
      {
        user: this.userId,
        file_id: fileId,
        bytes,
        filepath: filePath,
        filename,
        source: FileSources.local,
        type: mime,
        context: FileContext.message_attachment,
      },
      true,
    );

    const secret = process.env.JWT_SECRET;
    if (!secret) {
      throw new Error('JWT_SECRET is required to generate Office download links.');
    }
    const expires = Date.now() + 5 * 60 * 1000;
    const payload = `${this.userId}:${fileId}:${expires}`;
    const signature = crypto.createHmac('sha256', secret).update(payload).digest('hex');
    const downloadUrl =
      `/api/files/excel-download/${this.userId}/${fileId}` +
      `?expires=${expires}&signature=${signature}`;

    return {
      file_id: fileId,
      output_file: filename,
      download_url: downloadUrl,
      download_markdown: `[Descargar archivo ${label}](${downloadUrl})`,
    };
  }

  async createWord(data) {
    if (!data.title) {
      throw new Error('create_word requires title.');
    }
    const userDirectory = this.ensureUserDirectory();
    const filename = sanitizeFilename(data.file_name, '.docx', 'documento.docx');
    const filePath = path.join(userDirectory, `${uuidv4()}__${filename}`);
    const buffer = await buildDocx(data);
    fs.writeFileSync(filePath, buffer);
    const registered = await this.registerGeneratedFile(filePath, filename, WORD_MIME, 'Word');
    return {
      success: true,
      format: 'docx',
      title: data.title,
      sections: Array.isArray(data.sections) ? data.sections.length : 0,
      ...registered,
    };
  }

  async createPowerPoint(data) {
    const slides = Array.isArray(data.slides) ? data.slides : [];
    if (!slides.length && !data.title) {
      throw new Error('create_powerpoint requires slides or a title.');
    }
    const userDirectory = this.ensureUserDirectory();
    const filename = sanitizeFilename(data.file_name, '.pptx', 'presentacion.pptx');
    const filePath = path.join(userDirectory, `${uuidv4()}__${filename}`);
    const buffer = await buildPptx(data);
    fs.writeFileSync(filePath, buffer);
    const registered = await this.registerGeneratedFile(filePath, filename, PPT_MIME, 'PowerPoint');
    return {
      success: true,
      format: 'pptx',
      title: data.title || slides[0]?.title || 'Presentación',
      slides: slides.length || 1,
      ...registered,
    };
  }

  async _call(data) {
    try {
      if (data.action === 'create_word') {
        return JSON.stringify(await this.createWord(data));
      }
      if (data.action === 'create_powerpoint') {
        return JSON.stringify(await this.createPowerPoint(data));
      }
      return super._call(data);
    } catch (error) {
      return JSON.stringify({ success: false, error: error.message });
    }
  }
}

module.exports = OfficeTool;
