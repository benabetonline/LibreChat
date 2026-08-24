import type { ExtendedJsonSchema } from './schema';
import { toolDefinitions } from './definitions';

const wordToolSchema: ExtendedJsonSchema = {
  type: 'object',
  properties: {
    action: {
      type: 'string',
      enum: ['create_word'],
      description: 'Create a downloadable Microsoft Word DOCX file.',
    },
    file_name: {
      type: 'string',
      description: 'Output filename, for example informe.docx.',
    },
    title: {
      type: 'string',
      description: 'Document title.',
    },
    subtitle: {
      type: 'string',
      description: 'Optional document subtitle.',
    },
    author: {
      type: 'string',
      description: 'Optional author name.',
    },
    sections: {
      type: 'array',
      maxItems: 60,
      description: 'Complete document content. Include all requested sections here.',
      items: {
        type: 'object',
        properties: {
          heading: { type: 'string' },
          level: { type: 'number', minimum: 1, maximum: 3 },
          paragraphs: { type: 'array', items: { type: 'string' } },
          bullets: { type: 'array', items: { type: 'string' } },
          table: {
            type: 'object',
            properties: {
              headers: { type: 'array', items: { type: 'string' } },
              rows: {
                type: 'array',
                items: { type: 'array', items: { type: 'string' } },
              },
            },
          },
        },
      },
    },
  },
  required: ['action', 'title'],
};

const powerpointToolSchema: ExtendedJsonSchema = {
  type: 'object',
  properties: {
    action: {
      type: 'string',
      enum: ['create_powerpoint'],
      description: 'Create a downloadable Microsoft PowerPoint PPTX file.',
    },
    file_name: {
      type: 'string',
      description: 'Output filename, for example presentacion.pptx.',
    },
    title: {
      type: 'string',
      description: 'Presentation title.',
    },
    subtitle: {
      type: 'string',
      description: 'Optional presentation subtitle.',
    },
    author: {
      type: 'string',
      description: 'Optional author name.',
    },
    slides: {
      type: 'array',
      minItems: 1,
      maxItems: 30,
      description: 'Complete slide deck to generate.',
      items: {
        type: 'object',
        properties: {
          layout: { type: 'string', enum: ['title', 'content', 'section'] },
          title: { type: 'string' },
          subtitle: { type: 'string' },
          paragraphs: { type: 'array', items: { type: 'string' } },
          bullets: { type: 'array', maxItems: 12, items: { type: 'string' } },
        },
        required: ['title'],
      },
    },
  },
  required: ['action', 'slides'],
};

toolDefinitions.word_tool = {
  name: 'word_tool',
  description:
    'Create downloadable Microsoft Word (.docx) documents. Use this tool whenever the user asks for a Word or DOCX file.',
  description_for_model:
    'When the user requests a Word/DOCX document, call word_tool with action="create_word", provide the complete content in sections, and return the generated download link. Do not claim Word generation is unavailable.',
  schema: wordToolSchema,
  toolType: 'builtin',
};

toolDefinitions.powerpoint_tool = {
  name: 'powerpoint_tool',
  description:
    'Create downloadable Microsoft PowerPoint (.pptx) presentations. Use this tool whenever the user asks for a PowerPoint or PPTX file.',
  description_for_model:
    'When the user requests a PowerPoint/PPTX presentation, call powerpoint_tool with action="create_powerpoint", provide the complete slide deck, and return the generated download link. Do not claim PowerPoint generation is unavailable.',
  schema: powerpointToolSchema,
  toolType: 'builtin',
};
