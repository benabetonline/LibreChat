const OfficeTool = require('./OfficeTool');

const powerpointJsonSchema = {
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
  required: ['action', 'slides'],
};

class PowerPointTool extends OfficeTool {
  constructor(fields = {}) {
    super(fields);
    this.name = 'powerpoint_tool';
    this.description = `
Create downloadable Microsoft PowerPoint (.pptx) presentations for the current user.
Use this tool whenever the user asks for a PowerPoint or PPTX presentation.
Provide the complete slide deck and return the generated download link.
Do not tell the user that PowerPoint generation is unavailable when this tool is present.
`;
    this.schema = powerpointJsonSchema;
  }

  static get jsonSchema() {
    return powerpointJsonSchema;
  }

  async _call(data) {
    return super._call({ ...data, action: 'create_powerpoint' });
  }
}

module.exports = PowerPointTool;
