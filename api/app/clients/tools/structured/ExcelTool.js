const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const { Tool } = require('@librechat/agents/langchain/tools');

const excelJsonSchema = {
  type: 'object',
  properties: {
    action: {
      type: 'string',
      enum: ['inspect'],
      description: 'Action to perform on the Excel file. Currently supports inspect.',
    },
    file_name: {
      type: 'string',
      description:
        'Optional Excel filename. If omitted, the most recently uploaded Excel or CSV file for the current user will be used.',
    },
  },
  required: ['action'],
};

class ExcelTool extends Tool {
  constructor(fields = {}) {
    super();

    this.userId = fields.userId;
    this.name = 'excel_tool';

    this.description = `
Read and inspect Excel or CSV files uploaded by the current user.
Use this tool when the user asks about the contents, sheets, rows, columns,
or data contained in an Excel (.xlsx, .xls) or CSV file.
The tool can only access files belonging to the current LibreChat user.
`;

    this.schema = excelJsonSchema;
  }

  static get jsonSchema() {
    return excelJsonSchema;
  }

  getUserDirectory() {
    if (!this.userId) {
      throw new Error('ExcelTool: missing userId.');
    }

    const uploadsRoot = path.resolve('/app/uploads');
    const userDirectory = path.resolve(uploadsRoot, String(this.userId));

    if (!userDirectory.startsWith(`${uploadsRoot}${path.sep}`)) {
      throw new Error('ExcelTool: invalid user directory.');
    }

    if (!fs.existsSync(userDirectory)) {
      throw new Error('ExcelTool: user upload directory does not exist.');
    }

    return userDirectory;
  }

  findExcelFiles(directory) {
    const files = [];

    const walk = (currentDirectory) => {
      for (const entry of fs.readdirSync(currentDirectory, {
        withFileTypes: true,
      })) {
        const fullPath = path.join(currentDirectory, entry.name);

        if (entry.isDirectory()) {
          walk(fullPath);
          continue;
        }

        if (!/\.(xlsx|xls|csv)$/i.test(entry.name)) {
          continue;
        }

        const stats = fs.statSync(fullPath);

        files.push({
          path: fullPath,
          name: entry.name,
          modified: stats.mtimeMs,
        });
      }
    };

    walk(directory);

    return files.sort((a, b) => b.modified - a.modified);
  }

  selectFile(files, requestedName) {
    if (!files.length) {
      throw new Error('No Excel or CSV files were found for this user.');
    }

    if (!requestedName) {
      return files[0];
    }

    const requested = requestedName.toLowerCase();

    const match = files.find((file) =>
      file.name.toLowerCase().includes(requested),
    );

    if (!match) {
      throw new Error(`Excel file "${requestedName}" was not found.`);
    }

    return match;
  }

  inspectWorkbook(filePath) {
    const workbook = XLSX.readFile(filePath);

    const sheets = workbook.SheetNames.map((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];

      const rows = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: null,
      });

      return {
        name: sheetName,
        row_count: rows.length,
        preview: rows.slice(0, 20),
      };
    });

    return {
      sheets,
    };
  }

  async _call(data) {
    try {
      const userDirectory = this.getUserDirectory();
      const files = this.findExcelFiles(userDirectory);
      const selectedFile = this.selectFile(files, data.file_name);

      if (data.action === 'inspect') {
        const result = this.inspectWorkbook(selectedFile.path);

        return JSON.stringify({
          success: true,
          file: selectedFile.name,
          ...result,
        });
      }

      return JSON.stringify({
        success: false,
        error: 'Unsupported Excel action.',
      });
    } catch (error) {
      return JSON.stringify({
        success: false,
        error: error.message,
      });
    }
  }
}

module.exports = ExcelTool;
