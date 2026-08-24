const OfficeTool = require('./OfficeTool');

class OfficeDocumentsTool extends OfficeTool {
  constructor(fields = {}) {
    super(fields);
    this.name = 'office_documents';
    this.description = `
Microsoft Office document tool for the current LibreChat user.
You MUST use this tool whenever the user asks you to create a downloadable Microsoft Word (.docx) or PowerPoint (.pptx) file.
Use action="create_word" to create Word documents and action="create_powerpoint" to create PowerPoint presentations.
The same tool also supports the existing Excel/CSV actions: inspect, replace_text and batch_edit.
Do not tell the user that Word or PowerPoint generation is unavailable when this tool is present.
Generated files are private to the current user and receive short-lived signed download links.
`;
  }

  static get jsonSchema() {
    return OfficeTool.jsonSchema;
  }
}

module.exports = OfficeDocumentsTool;
