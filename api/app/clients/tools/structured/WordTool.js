const OfficeTool = require('./OfficeTool');

const wordJsonSchema = {
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
      description: 'Complete content of the Word document.',
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

class WordTool extends OfficeTool {
  constructor(fields = {}) {
    super(fields);
    this.name = 'word_tool';
    this.description = `
Create downloadable Microsoft Word (.docx) documents for the current user.
Use this tool whenever the user asks for a Word document or DOCX file.
Build the requested document in sections and return the generated download link.
Do not tell the user that Word generation is unavailable when this tool is present.
`;
    this.schema = wordJsonSchema;
  }

  static get jsonSchema() {
    return wordJsonSchema;
  }

  async _call(data) {
    return super._call({ ...data, action: 'create_word' });
  }
}

module.exports = WordTool;
