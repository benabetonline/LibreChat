import type { ExtendedJsonSchema } from './schema';
import { toolDefinitions } from './definitions';

/**
 * Keep the persisted tool key `excel_tool` for backwards compatibility with
 * existing agents, but expose the full Office feature set to the model.
 */
const officeToolSchema: ExtendedJsonSchema = {
  type: 'object',
  properties: {
    action: {
      type: 'string',
      enum: ['inspect', 'replace_text', 'batch_edit', 'create_word', 'create_powerpoint'],
      description:
        'Office action. Use inspect, replace_text or batch_edit for Excel/CSV; create_word to generate a downloadable DOCX; create_powerpoint to generate a downloadable PPTX.',
    },
    file_name: {
      type: 'string',
      description:
        'For Excel actions, optional uploaded filename. For Word or PowerPoint creation, optional output filename such as informe.docx or presentacion.pptx.',
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
      description: 'Optional Excel worksheet name.',
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
      description: 'Required title for a generated Word document or recommended title for a PowerPoint.',
    },
    subtitle: {
      type: 'string',
      description: 'Optional document or presentation subtitle.',
    },
    author: {
      type: 'string',
      description: 'Optional author name.',
    },
    sections: {
      type: 'array',
      maxItems: 60,
      description:
        'Content for create_word. Build the complete requested document here; do not return only prose when the user asks for a downloadable Word file.',
      items: {
        type: 'object',
        properties: {
          heading: { type: 'string' },
          level: { type: 'number', minimum: 1, maximum: 3 },
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
        'Complete slide deck for create_powerpoint. Use this when the user requests a downloadable PowerPoint presentation.',
      items: {
        type: 'object',
        properties: {
          layout: {
            type: 'string',
            enum: ['title', 'content', 'section'],
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

toolDefinitions.excel_tool = {
  name: 'excel_tool',
  description:
    'Microsoft Office tool. It can inspect and edit Excel/CSV files AND generate downloadable Microsoft Word (.docx) and PowerPoint (.pptx) files. When the user asks for a Word file, use action="create_word". When the user asks for a PowerPoint file, use action="create_powerpoint". Do not tell the user these formats cannot be generated.',
  description_for_model:
    'Use this Office tool for Excel work and whenever a user asks to create a downloadable Word DOCX or PowerPoint PPTX. For DOCX call create_word with title and sections. For PPTX call create_powerpoint with slides.',
  schema: officeToolSchema,
  toolType: 'builtin',
};
