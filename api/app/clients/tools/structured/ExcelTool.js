const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const XLSX = require('xlsx');
const { v4: uuidv4 } = require('uuid');
const { FileContext, FileSources } = require('librechat-data-provider');
const db = require('~/models');

const { Tool } = require('@librechat/agents/langchain/tools');

const excelJsonSchema = {
  type: 'object',
  properties: {
    action: {
      type: 'string',
      enum: ['inspect', 'replace_text', 'batch_edit'],
      description:
  'Action to perform on the Excel file. Supports inspect, replace_text and batch_edit.',
    },
    file_name: {
      type: 'string',
      description:
        'Optional Excel filename. If omitted, the most recently uploaded Excel or CSV file for the current user will be used.',
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
    'Optional worksheet name. If omitted, the replacement will be applied to all worksheets.',
},
column: {
  type: 'string',
  description:
    'Optional Excel column letter such as F or H. If omitted, all columns will be searched.',
},
    operations: {
  type: 'array',
  description:
    'List of Excel operations to execute in order when action is batch_edit.',
  items: {
    type: 'object',
    properties: {
      type: {
        type: 'string',
        enum: ['replace_text', 'set_cell', 'clear_cell', 'count_value'],
        description: 'Type of operation to execute.',
      },
      sheet_name: {
        type: 'string',
        description: 'Worksheet where the operation will be applied.',
      },
      column: {
        type: 'string',
        description:
          'Optional Excel column letter such as F. Used by replace_text or count_value.',
      },
      search_text: {
        type: 'string',
        description:
          'Value to search for when using replace_text or count_value.',
      },
      replacement_text: {
        type: 'string',
        description: 'Replacement value when using replace_text.',
      },
      cell: {
        type: 'string',
        description:
          'Target cell such as H11. Used by set_cell or clear_cell.',
      },
      value: {
        type: 'string',
        description: 'Value to write when using set_cell.',
      },
      output_cell: {
        type: 'string',
        description:
          'Cell where the numeric result of count_value will be written.',
      },
    },
    required: ['type'],
  },
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
async replaceText(filePath, data) {
  if (data.search_text === undefined || data.search_text === null) {
    throw new Error('replace_text requires search_text.');
  }

  if (data.replacement_text === undefined || data.replacement_text === null) {
    throw new Error('replace_text requires replacement_text.');
  }

  const workbook = XLSX.readFile(filePath);

  let sheetNames = workbook.SheetNames;

  if (data.sheet_name) {
    if (!workbook.SheetNames.includes(data.sheet_name)) {
      throw new Error(`Worksheet "${data.sheet_name}" was not found.`);
    }

    sheetNames = [data.sheet_name];
  }

  const requestedColumn = data.column
    ? String(data.column).trim().toUpperCase()
    : null;

  if (requestedColumn && !/^[A-Z]{1,3}$/.test(requestedColumn)) {
    throw new Error('Invalid Excel column. Use a letter such as F or H.');
  }

  const searchValue = String(data.search_text).trim();
  const replacementValue = String(data.replacement_text);

  let replacements = 0;
  const changedSheets = [];

  for (const sheetName of sheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    let sheetReplacements = 0;

    for (const address of Object.keys(worksheet)) {
      if (address.startsWith('!')) {
        continue;
      }

      const cell = worksheet[address];

      if (!cell || cell.f || cell.v === undefined || cell.v === null) {
        continue;
      }

      if (requestedColumn) {
        const decoded = XLSX.utils.decode_cell(address);
        const cellColumn = XLSX.utils.encode_col(decoded.c);

        if (cellColumn !== requestedColumn) {
          continue;
        }
      }

      if (String(cell.v).trim() === searchValue) {
        cell.v = replacementValue;
        cell.t = 's';

        replacements += 1;
        sheetReplacements += 1;
      }
    }

    if (sheetReplacements > 0) {
      changedSheets.push({
        name: sheetName,
        replacements: sheetReplacements,
      });
    }
  }

  const extension = path.extname(filePath);
  const baseName = path.basename(filePath, extension);
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');

  const outputPath = path.join(
    path.dirname(filePath),
    `${baseName}_modificado_${timestamp}.xlsx`,
  );

  XLSX.writeFile(workbook, outputPath);
const outputFile = path.basename(outputPath);
const fileId = uuidv4();
const bytes = fs.statSync(outputPath).size;

await db.createFile(
  {
    user: this.userId,
    file_id: fileId,
    bytes,
    filepath: outputPath,
    filename: outputFile,
    source: FileSources.local,
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    context: FileContext.message_attachment,
  },
  true,
);

const secret = process.env.JWT_SECRET;

if (!secret) {
  throw new Error('JWT_SECRET is required to generate Excel download links.');
}

const expires = Date.now() + 5 * 60 * 1000;
const payload = `${this.userId}:${fileId}:${expires}`;

const signature = crypto
  .createHmac('sha256', secret)
  .update(payload)
  .digest('hex');

const downloadUrl =
  `/api/files/excel-download/${this.userId}/${fileId}` +
  `?expires=${expires}&signature=${signature}`;
  
  return {
  output_path: outputPath,
  output_file: outputFile,
  file_id: fileId,
  download_url: downloadUrl,
  download_markdown: `[Descargar archivo Excel](${downloadUrl})`,
  replacements,
  changed_sheets: changedSheets,
};
}
  async batchEdit(filePath, data) {
  if (!Array.isArray(data.operations) || data.operations.length === 0) {
    throw new Error('batch_edit requires at least one operation.');
  }

  const workbook = XLSX.readFile(filePath);
  const results = [];

  const getSheet = (sheetName) => {
    const targetName = sheetName || workbook.SheetNames[0];

    if (!workbook.SheetNames.includes(targetName)) {
      throw new Error(`Worksheet "${targetName}" was not found.`);
    }

    return {
      name: targetName,
      worksheet: workbook.Sheets[targetName],
    };
  };

  const validateColumn = (column) => {
    if (!column) {
      return null;
    }

    const value = String(column).trim().toUpperCase();

    if (!/^[A-Z]{1,3}$/.test(value)) {
      throw new Error(`Invalid Excel column "${column}".`);
    }

    return value;
  };

  const validateCell = (cell) => {
    const value = String(cell || '').trim().toUpperCase();

    if (!/^[A-Z]{1,3}[1-9][0-9]*$/.test(value)) {
      throw new Error(`Invalid Excel cell "${cell}".`);
    }

    return value;
  };

  for (const operation of data.operations) {
    if (operation.type === 'replace_text') {
      if (
        operation.search_text === undefined ||
        operation.replacement_text === undefined
      ) {
        throw new Error(
          'replace_text requires search_text and replacement_text.',
        );
      }

      const requestedColumn = validateColumn(operation.column);

      const sheetNames = operation.sheet_name
        ? [operation.sheet_name]
        : workbook.SheetNames;

      let replacements = 0;

      for (const sheetName of sheetNames) {
        if (!workbook.SheetNames.includes(sheetName)) {
          throw new Error(`Worksheet "${sheetName}" was not found.`);
        }

        const worksheet = workbook.Sheets[sheetName];

        for (const address of Object.keys(worksheet)) {
          if (address.startsWith('!')) {
            continue;
          }

          const cell = worksheet[address];

          if (!cell || cell.f || cell.v === undefined || cell.v === null) {
            continue;
          }

          if (requestedColumn) {
            const decoded = XLSX.utils.decode_cell(address);
            const cellColumn = XLSX.utils.encode_col(decoded.c);

            if (cellColumn !== requestedColumn) {
              continue;
            }
          }

          if (
            String(cell.v).trim() ===
            String(operation.search_text).trim()
          ) {
            cell.v = String(operation.replacement_text);
            cell.t = 's';
            replacements += 1;
          }
        }
      }

      results.push({
        type: 'replace_text',
        search_text: operation.search_text,
        replacement_text: operation.replacement_text,
        replacements,
      });

      continue;
    }

    if (operation.type === 'set_cell') {
      const { name, worksheet } = getSheet(operation.sheet_name);
      const address = validateCell(operation.cell);

      worksheet[address] = {
        ...(worksheet[address] || {}),
        v: operation.value ?? '',
        t: typeof operation.value === 'number' ? 'n' : 's',
      };

      delete worksheet[address].f;
      delete worksheet[address].w;

      results.push({
        type: 'set_cell',
        sheet: name,
        cell: address,
        value: operation.value ?? '',
      });

      continue;
    }

    if (operation.type === 'clear_cell') {
      const { name, worksheet } = getSheet(operation.sheet_name);
      const address = validateCell(operation.cell);

      if (worksheet[address]) {
        worksheet[address].v = '';
        worksheet[address].t = 's';
        delete worksheet[address].f;
        delete worksheet[address].w;
      }

      results.push({
        type: 'clear_cell',
        sheet: name,
        cell: address,
      });

      continue;
    }

    if (operation.type === 'count_value') {
      if (operation.search_text === undefined) {
        throw new Error('count_value requires search_text.');
      }

      const { name, worksheet } = getSheet(operation.sheet_name);
      const requestedColumn = validateColumn(operation.column);

      let count = 0;

      for (const address of Object.keys(worksheet)) {
        if (address.startsWith('!')) {
          continue;
        }

        const cell = worksheet[address];

        if (!cell || cell.v === undefined || cell.v === null) {
          continue;
        }

        if (requestedColumn) {
          const decoded = XLSX.utils.decode_cell(address);
          const cellColumn = XLSX.utils.encode_col(decoded.c);

          if (cellColumn !== requestedColumn) {
            continue;
          }
        }

        if (
          String(cell.v).trim() ===
          String(operation.search_text).trim()
        ) {
          count += 1;
        }
      }

      if (operation.output_cell) {
        const outputAddress = validateCell(operation.output_cell);

        worksheet[outputAddress] = {
          ...(worksheet[outputAddress] || {}),
          v: count,
          t: 'n',
        };

        delete worksheet[outputAddress].f;
        delete worksheet[outputAddress].w;
      }

      results.push({
        type: 'count_value',
        sheet: name,
        search_text: operation.search_text,
        count,
        output_cell: operation.output_cell || null,
      });

      continue;
    }

    throw new Error(`Unsupported batch operation "${operation.type}".`);
  }

  const extension = path.extname(filePath);
  const baseName = path.basename(filePath, extension);
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');

  const outputPath = path.join(
    path.dirname(filePath),
    `${baseName}_batch_${timestamp}.xlsx`,
  );

  XLSX.writeFile(workbook, outputPath);

  const outputFile = path.basename(outputPath);
  const fileId = uuidv4();
  const bytes = fs.statSync(outputPath).size;

  await db.createFile(
    {
      user: this.userId,
      file_id: fileId,
      bytes,
      filepath: outputPath,
      filename: outputFile,
      source: FileSources.local,
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      context: FileContext.message_attachment,
    },
    true,
  );

  const secret = process.env.JWT_SECRET;

  if (!secret) {
    throw new Error(
      'JWT_SECRET is required to generate Excel download links.',
    );
  }

  const expires = Date.now() + 5 * 60 * 1000;
  const payload = `${this.userId}:${fileId}:${expires}`;

  const signature = crypto
    .createHmac('sha256', secret)
    .update(payload)
    .digest('hex');

  const downloadUrl =
    `/api/files/excel-download/${this.userId}/${fileId}` +
    `?expires=${expires}&signature=${signature}`;

  return {
    output_file: outputFile,
    file_id: fileId,
    download_url: downloadUrl,
    download_markdown: `[Descargar archivo Excel](${downloadUrl})`,
    operations_executed: results,
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
if (data.action === 'replace_text') {
  const result = await this.replaceText(selectedFile.path, data);

  return JSON.stringify({
    success: true,
    original_file: selectedFile.name,
    ...result,
  });
}
      if (data.action === 'batch_edit') {
  const result = await this.batchEdit(selectedFile.path, data);

  return JSON.stringify({
    success: true,
    original_file: selectedFile.name,
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
