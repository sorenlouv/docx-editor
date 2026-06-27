/**
 * Tables (`w:tbl`), rows (`w:tr`), and cells (`w:tc`).
 */

import type { TableFormatting, TableRowFormatting, TableCellFormatting } from '../formatting';
import type { Paragraph } from './paragraph';
import type { BookmarkStart, BookmarkEnd } from './link';
import type {
  TablePropertyChange,
  TableRowPropertyChange,
  TableCellPropertyChange,
  TableStructuralChangeInfo,
} from './trackedChange';

/**
 * Table cell (`w:tc`). Holds nested block content (paragraphs and nested
 * tables), cell-level formatting (borders, shading, vertical merge),
 * tracked property changes, and tracked structural changes for cell
 * insert/delete/merge operations.
 */
export interface TableCell {
  type: 'tableCell';
  /** Cell formatting */
  formatting?: TableCellFormatting;
  /** Cell-level tracked property changes (w:tcPrChange) */
  propertyChanges?: TableCellPropertyChange[];
  /** Tracked structural changes (cell insert/delete/merge) */
  structuralChange?: TableStructuralChangeInfo;
  /** Cell content (paragraphs, tables, etc.) */
  content: (Paragraph | Table)[];
}

/**
 * Table row (`w:tr`) — an ordered list of `TableCell` plus row-level
 * formatting (height, repeated header, cantSplit) and tracked changes
 * for inserts/deletes.
 */
export interface TableRow {
  type: 'tableRow';
  /** Row formatting */
  formatting?: TableRowFormatting;
  /** Row-level tracked property changes (w:trPrChange) */
  propertyChanges?: TableRowPropertyChange[];
  /** Tracked structural changes (row insert/delete) */
  structuralChange?: TableStructuralChangeInfo;
  /** Cells in this row */
  cells: TableCell[];
}

/**
 * Table (`w:tbl`) — a block-level grid of rows × cells. Tables carry
 * their own formatting layer (borders, shading, alignment, indent,
 * floating placement) and an explicit column-width grid in twips. Tables
 * can nest arbitrarily through `TableCell.content`.
 *
 * See ECMA-376 §17.4.
 */
export interface Table {
  type: 'table';
  /** Table formatting */
  formatting?: TableFormatting;
  /** Table-level tracked property changes (w:tblPrChange) */
  propertyChanges?: TablePropertyChange[];
  /** Column widths in twips */
  columnWidths?: number[];
  /** Table rows */
  rows: TableRow[];
  /**
   * Block-level bookmark markers that sit as direct children of the parent
   * block container immediately BEFORE this table's `w:tbl`, i.e.
   * `<w:bookmarkStart/><w:tbl>`. See {@link Paragraph.leadingBlockMarkers}.
   */
  leadingBlockMarkers?: (BookmarkStart | BookmarkEnd)[];
  /**
   * Block-level bookmark markers that sit immediately AFTER this table's
   * `w:tbl` (e.g. `<w:tbl></w:tbl><w:bookmarkEnd/>`). See
   * {@link Paragraph.leadingBlockMarkers}.
   */
  trailingBlockMarkers?: (BookmarkStart | BookmarkEnd)[];
}
