// ============================================
// XLSX Stream Writer — Bun-native disk-backed
// streaming via FileSink/temp files
// ============================================

import { resolve } from 'node:path';
import type {
  Cell,
  CellRange,
  CellStyle,
  CellValue,
  ColumnConfig,
  ConditionalFormatting,
  DataValidation,
  ExcelWriteOptions,
  MergeCell,
  Row,
  StreamWriter,
  Workbook,
  Worksheet,
} from '../types';
import { ExcelChunkedStreamWriter } from './xlsx-chunked-stream-writer';

/** Validate path for security */
function validatePath(filePath: string): string {
  if (filePath.includes('\0')) {
    throw new Error('Invalid file path: contains null bytes');
  }
  return resolve(filePath);
}

/**
 * Options for the Excel stream writer
 */
export interface ExcelStreamOptions extends ExcelWriteOptions {
  /** Sheet name (default: "Sheet1") */
  sheetName?: string;
  /** Column configurations */
  columns?: ColumnConfig[];
  /** Default row height */
  defaultRowHeight?: number;
  /** Freeze pane */
  freezePane?: { row: number; col: number };
  /** Split pane */
  splitPane?: Worksheet['splitPane'];
  /** Merge cells */
  mergeCells?: MergeCell[];
  /** Auto filter range */
  autoFilter?: CellRange;
  /** Conditional formatting rules */
  conditionalFormattings?: ConditionalFormatting[];
  /** Data validation rules */
  dataValidations?: DataValidation[];
}

/**
 * Excel Stream Writer — Bun-native disk-backed streaming
 *
 * Delegates to the disk-backed chunked writer so the public
 * createExcelStream() API also uses Bun FileSink/temp files
 * instead of keeping row XML in memory.
 */
export class ExcelStreamWriter implements StreamWriter {
  private readonly writer: ExcelChunkedStreamWriter;

  constructor(path: string, options?: ExcelStreamOptions) {
    this.writer = new ExcelChunkedStreamWriter(validatePath(path), options);
  }

  /**
   * Write a single row
   */
  writeRow(row: Row | CellValue[]): void {
    this.writer.writeRow(row);
  }

  /**
   * Write a row with styles applied to each cell
   */
  writeStyledRow(values: CellValue[], styles: (CellStyle | undefined)[]): void {
    this.writer.writeStyledRow(values, styles);
  }

  /**
   * Write multiple rows at once
   */
  writeRows(rows: (Row | CellValue[])[]): void {
    this.writer.writeRows(rows);
  }

  /**
   * Flush buffered temp-file writes.
   */
  flush(): Promise<void> {
    return this.writer.flush();
  }

  /**
   * Finalize and write the XLSX file.
   */
  async end(): Promise<void> {
    await this.writer.end();
  }

  /**
   * Get current row count
   */
  get currentRowCount(): number {
    return this.writer.currentRowCount;
  }
}

/**
 * Multi-sheet Excel Stream Writer
 * Allows streaming data to multiple worksheets
 */
export class MultiSheetExcelStreamWriter {
  private worksheets: Map<string, { rows: Row[]; config: ExcelStreamOptions }> =
    new Map();
  private path: string;
  private options: ExcelWriteOptions;
  private currentSheet: string;

  constructor(path: string, options?: ExcelWriteOptions) {
    this.path = validatePath(path);
    this.options = options || {};
    this.currentSheet = 'Sheet1';
    this.worksheets.set('Sheet1', { rows: [], config: {} });
  }

  /**
   * Add a new sheet or switch to existing sheet
   */
  addSheet(name: string, config?: ExcelStreamOptions): this {
    if (!this.worksheets.has(name)) {
      this.worksheets.set(name, { rows: [], config: config || {} });
    }
    this.currentSheet = name;
    return this;
  }

  /**
   * Write a row to the current sheet
   */
  writeRow(row: Row | CellValue[]): void {
    const sheet = this.worksheets.get(this.currentSheet);
    if (!sheet) throw new Error(`Sheet not found: ${this.currentSheet}`);

    if (Array.isArray(row)) {
      sheet.rows.push({ cells: row.map((value) => ({ value })) });
    } else {
      sheet.rows.push(row);
    }
  }

  /**
   * Write a styled row to the current sheet
   */
  writeStyledRow(values: CellValue[], styles: (CellStyle | undefined)[]): void {
    const sheet = this.worksheets.get(this.currentSheet);
    if (!sheet) throw new Error(`Sheet not found: ${this.currentSheet}`);

    const cells: Cell[] = values.map((value, i) => ({
      value,
      style: styles[i],
    }));
    sheet.rows.push({ cells });
  }

  /**
   * Finalize and write the Excel file
   */
  async end(): Promise<void> {
    const worksheetsList: Worksheet[] = [];

    for (const [name, data] of this.worksheets) {
      worksheetsList.push({
        name,
        rows: data.rows,
        columns: data.config.columns,
        autoFilter: data.config.autoFilter,
        conditionalFormattings: data.config.conditionalFormattings,
        dataValidations: data.config.dataValidations,
        freezePane: data.config.freezePane,
        splitPane: data.config.splitPane,
        defaultRowHeight: data.config.defaultRowHeight,
      });
    }

    const workbook: Workbook = {
      worksheets: worksheetsList,
      creator: this.options.creator,
      created: this.options.created,
      modified: this.options.modified,
    };

    const { writeExcel } = await import('./xlsx-writer');
    await writeExcel(this.path, workbook, this.options);

    // Clear data
    this.worksheets.clear();
  }
}

/**
 * Create an Excel stream writer (disk-backed Bun-native streaming)
 */
export function createExcelStream(
  path: string,
  options?: ExcelStreamOptions,
): ExcelStreamWriter {
  return new ExcelStreamWriter(path, options);
}

/**
 * Create a multi-sheet Excel stream writer
 */
export function createMultiSheetExcelStream(
  path: string,
  options?: ExcelWriteOptions,
): MultiSheetExcelStreamWriter {
  return new MultiSheetExcelStreamWriter(path, options);
}
