import type { Row, Worksheet } from '../types';
import {
  buildCommentsVML,
  buildCommentsXML,
  type CommentEntry,
  commentRefFromCoords,
} from './comments';
import { buildDrawingArtifacts, type DrawingArtifacts } from './images';
import { buildTableXML } from './tables';

export interface SheetRelationship {
  id: string;
  type: string;
  target: string;
  targetMode?: 'External';
}

export interface SheetExtraFile {
  path: string;
  content: Uint8Array;
}

export interface WorksheetFeatureArtifacts {
  relationships: SheetRelationship[];
  xmlPartsBeforeClose: string[];
  extraFiles: SheetExtraFile[];
  mediaExtensions: Set<string>;
  commentCount: number;
  drawingCount: number;
  tableCount: number;
}

const encoder = new TextEncoder();

export function buildSheetRelsXML(relationships: SheetRelationship[]): string {
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  xml +=
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
  for (const relationship of relationships) {
    xml += `<Relationship Id="${relationship.id}" Type="${relationship.type}" Target="${relationship.target}"`;
    if (relationship.targetMode) {
      xml += ` TargetMode="${relationship.targetMode}"`;
    }
    xml += '/>';
  }
  xml += '</Relationships>';
  return xml;
}

export function collectCommentEntries(rows: Row[]): CommentEntry[] {
  const entries: CommentEntry[] = [];
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    for (let c = 0; c < row.cells.length; c++) {
      const cell = row.cells[c];
      if (!cell?.comment) continue;
      entries.push({
        ref: commentRefFromCoords(r, c),
        comment: cell.comment,
      });
    }
  }
  return entries;
}

export function buildWorksheetFeatureArtifacts(
  worksheet: Worksheet,
  counters: {
    nextCommentsIndex: number;
    nextDrawingIndex: number;
    nextTableIndex: number;
  },
  options?: {
    commentEntries?: CommentEntry[];
    startingRelIndex?: number;
  },
): WorksheetFeatureArtifacts {
  const relationships: SheetRelationship[] = [];
  const xmlPartsBeforeClose: string[] = [];
  const extraFiles: SheetExtraFile[] = [];
  const mediaExtensions = new Set<string>();

  let relCounter = options?.startingRelIndex ?? 1;
  const nextRelId = () => `rId${relCounter++}`;

  const commentEntries =
    options?.commentEntries ?? collectCommentEntries(worksheet.rows);
  let commentCount = 0;
  if (commentEntries.length > 0) {
    const commentsIndex = counters.nextCommentsIndex++;
    const commentsRelId = nextRelId();
    const vmlRelId = nextRelId();
    relationships.push({
      id: commentsRelId,
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
      target: `../comments${commentsIndex}.xml`,
    });
    relationships.push({
      id: vmlRelId,
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
      target: `../drawings/vmlDrawing${commentsIndex}.vml`,
    });
    xmlPartsBeforeClose.push(`<legacyDrawing r:id="${vmlRelId}"/>`);
    extraFiles.push({
      path: `xl/comments${commentsIndex}.xml`,
      content: encoder.encode(buildCommentsXML(commentEntries)),
    });
    extraFiles.push({
      path: `xl/drawings/vmlDrawing${commentsIndex}.vml`,
      content: encoder.encode(buildCommentsVML(commentEntries)),
    });
    commentCount = 1;
  }

  let drawingCount = 0;
  if (worksheet.images && worksheet.images.length > 0) {
    const drawingIndex = counters.nextDrawingIndex++;
    const drawingRelId = nextRelId();
    const drawingArtifacts: DrawingArtifacts = buildDrawingArtifacts(
      drawingIndex,
      worksheet.images,
    );
    relationships.push({
      id: drawingRelId,
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
      target: `../drawings/drawing${drawingIndex}.xml`,
    });
    xmlPartsBeforeClose.push(`<drawing r:id="${drawingRelId}"/>`);
    extraFiles.push({
      path: `xl/drawings/drawing${drawingIndex}.xml`,
      content: encoder.encode(drawingArtifacts.drawingXml),
    });
    extraFiles.push({
      path: `xl/drawings/_rels/drawing${drawingIndex}.xml.rels`,
      content: encoder.encode(drawingArtifacts.drawingRelsXml),
    });
    for (const media of drawingArtifacts.media) {
      extraFiles.push({ path: media.path, content: media.data });
    }
    for (const image of worksheet.images) {
      mediaExtensions.add(image.format === 'jpg' ? 'jpeg' : image.format);
    }
    drawingCount = 1;
  }

  let tableCount = 0;
  if (worksheet.tables && worksheet.tables.length > 0) {
    const tableParts: string[] = [];
    for (const table of worksheet.tables) {
      const tableIndex = counters.nextTableIndex++;
      const tableRelId = nextRelId();
      relationships.push({
        id: tableRelId,
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
        target: `../tables/table${tableIndex}.xml`,
      });
      tableParts.push(`<tablePart r:id="${tableRelId}"/>`);
      extraFiles.push({
        path: `xl/tables/table${tableIndex}.xml`,
        content: encoder.encode(buildTableXML(worksheet, table, tableIndex)),
      });
      tableCount++;
    }
    xmlPartsBeforeClose.push(
      `<tableParts count="${tableCount}">${tableParts.join('')}</tableParts>`,
    );
  }

  return {
    relationships,
    xmlPartsBeforeClose,
    extraFiles,
    mediaExtensions,
    commentCount,
    drawingCount,
    tableCount,
  };
}
