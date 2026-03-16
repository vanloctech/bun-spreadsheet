import type { Cell, Row, Worksheet } from './types';

function cloneCell(cell: Cell): Cell {
  return {
    ...cell,
    style: cell.style ? { ...cell.style } : undefined,
    hyperlink: cell.hyperlink ? { ...cell.hyperlink } : undefined,
    comment: cell.comment ? { ...cell.comment } : undefined,
    richText: cell.richText?.map((run) => ({
      ...run,
      font: run.font ? { ...run.font } : undefined,
    })),
  };
}

function cloneRow(row: Row): Row {
  return {
    ...row,
    style: row.style ? { ...row.style } : undefined,
    cells: row.cells.map(cloneCell),
  };
}

function ensureColumnCount(worksheet: Worksheet, count: number) {
  if (!worksheet.columns) worksheet.columns = [];
  while (worksheet.columns.length < count) {
    worksheet.columns.push({});
  }
}

export function insertRows(
  worksheet: Worksheet,
  index: number,
  rows: Row[],
): Worksheet {
  const insertAt = Math.max(0, Math.min(index, worksheet.rows.length));
  worksheet.rows.splice(insertAt, 0, ...rows.map(cloneRow));
  return worksheet;
}

export function spliceRows(
  worksheet: Worksheet,
  index: number,
  deleteCount: number,
  rows: Row[] = [],
): Worksheet {
  worksheet.rows.splice(
    Math.max(0, index),
    Math.max(0, deleteCount),
    ...rows.map(cloneRow),
  );
  return worksheet;
}

export function duplicateRow(
  worksheet: Worksheet,
  rowIndex: number,
  count = 1,
): Worksheet {
  const row = worksheet.rows[rowIndex];
  if (!row || count <= 0) return worksheet;
  const clones = Array.from({ length: count }, () => cloneRow(row));
  worksheet.rows.splice(rowIndex + 1, 0, ...clones);
  return worksheet;
}

export function insertColumns(
  worksheet: Worksheet,
  index: number,
  count = 1,
): Worksheet {
  if (count <= 0) return worksheet;
  const insertAt = Math.max(0, index);
  ensureColumnCount(worksheet, insertAt);

  if (worksheet.columns) {
    worksheet.columns.splice(
      insertAt,
      0,
      ...Array.from({ length: count }, () => ({})),
    );
  }

  for (const row of worksheet.rows) {
    row.cells.splice(
      insertAt,
      0,
      ...Array.from({ length: count }, () => ({ value: null })),
    );
  }

  return worksheet;
}
