/**
 * @fileoverview CellWeave – Worksheet Model (Single Responsibility / Open-Closed)
 * Holds all data for a single worksheet. New features (charts, tables, CF…)
 * are added via separate registries rather than modifying this class.
 */

'use strict';

const { Cell } = require('./Cell');
const {
  decodeCell,
  decodeRange,
  encodeRange,
  encodeCell,
  sanitizeSheetName,
  rangesOverlap,
} = require('../utils/addressUtils');

/**
 * @class Worksheet
 * Stores cells, row/col config, merges, charts, tables, and all sheet-level features.
 */
class Worksheet {
  /**
   * @param {string} name
   * @param {number} [sheetId=1]
   */
  constructor(name, sheetId = 1) {
    this.name       = sanitizeSheetName(name);
    this.sheetId    = sheetId;

    /** @type {Map<string, Cell>}  key = "row:col" */
    this._cells = new Map();

    /** @type {Map<number, import('../interfaces').IRowProperties>} */
    this.rowProps = new Map();

    /** @type {Map<number, import('../interfaces').IColumnProperties>} */
    this.colProps = new Map();

    /** @type {Array<import('../interfaces').ICellRange>} merged cell ranges */
    this.merges = [];

    /** @type {Array<import('../interfaces').IConditionalFormatting>} */
    this.conditionalFormattings = [];

    /** @type {Array<import('../interfaces').IDataValidation>} */
    this.dataValidations = [];

    /** @type {import('../interfaces').IAutoFilter|null} */
    this.autoFilter = null;

    /** @type {import('../interfaces').ISortState|null} */
    this.sortState = null;

    /** @type {Array<import('../interfaces').ITable>} */
    this.tables = [];

    /** @type {Array<import('../interfaces').IChartConfig>} */
    this.charts = [];

    /** @type {Array<import('../interfaces').IImage>} */
    this.images = [];

    /** @type {Array<import('../interfaces').ISparkline>} */
    this.sparklines = [];

    /** @type {Array<import('../interfaces').ISheetGrouping>} */
    this.groupings = [];

    /** @type {Array<{row:number, col:number, rowSpan:number, colSpan:number}>} page breaks */
    this.rowPageBreaks = [];
    this.colPageBreaks = [];

    /** @type {import('../interfaces').IPageSetup} */
    this.pageSetup = {};

    /** @type {import('../interfaces').IPageMargins} */
    this.pageMargins = { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7, header: 0.3, footer: 0.3 };

    /** @type {import('../interfaces').IHeaderFooter} */
    this.headerFooter = {};

    /** @type {import('../interfaces').ISheetView} */
    this.sheetView = { showGridLines: true, showRowColHeaders: true, zoomScale: 100 };

    /** @type {import('../interfaces').IProtection|null} */
    this.protection = null;

    /** @type {string|null} tab color hex */
    this.tabColor = null;

    /** @type {number} maximum row index written (0-based) */
    this._maxRow = -1;
    /** @type {number} maximum col index written (0-based) */
    this._maxCol = -1;
  }

  // ── Cell Access ───────────────────────────────────────────────────────────

  /**
   * Get a cell by A1 address or 0-based indices.
   * @param {string|number} rowOrRef - row index or A1 address
   * @param {number} [col]
   * @returns {Cell|undefined}
   */
  getCell(rowOrRef, col) {
    const { row, c } = this._resolveRef(rowOrRef, col);
    return this._cells.get(`${row}:${c}`);
  }

  /**
   * Get or create a cell.
   * @param {string|number} rowOrRef
   * @param {number} [col]
   * @returns {Cell}
   */
  getCellOrCreate(rowOrRef, col) {
    const { row, c } = this._resolveRef(rowOrRef, col);
    const key = `${row}:${c}`;
    if (!this._cells.has(key)) {
      this._cells.set(key, new Cell(row, c));
    }
    return this._cells.get(key);
  }

  /**
   * Set cell value(s). Supports primitives, Date, formula objects.
   * @param {string|number} rowOrRef
   * @param {number|import('../interfaces').ICellData} colOrData
   * @param {import('../interfaces').ICellData} [data]
   * @returns {Cell}
   */
  setCell(rowOrRef, colOrData, data) {
    let row, c, cellData;
    if (typeof rowOrRef === 'string' && typeof colOrData !== 'number') {
      const decoded = decodeCell(rowOrRef);
      row = decoded.row; c = decoded.col; cellData = colOrData;
    } else if (typeof colOrData === 'number') {
      row = rowOrRef; c = colOrData; cellData = data || {};
    } else {
      throw new TypeError('Invalid setCell arguments');
    }

    const key = `${row}:${c}`;
    const existing = this._cells.get(key);
    const cell = new Cell(row, c, existing ? { ...existing.toJSON(), ...cellData } : cellData);
    this._cells.set(key, cell);
    this._trackBounds(row, c);
    return cell;
  }

  /**
   * Write a 2D array of values starting at (startRow, startCol).
   * @param {number} startRow
   * @param {number} startCol
   * @param {Array<Array<*>>} data
   * @param {import('../interfaces').ICellStyle} [style]
   */
  writeArray(startRow, startCol, data, style) {
    for (let r = 0; r < data.length; r++) {
      const row = data[r];
      for (let c = 0; c < row.length; c++) {
        const val = row[c];
        const cellData = (val !== null && typeof val === 'object' && !Array.isArray(val) && !(val instanceof Date))
          ? val
          : { value: val };
        if (style && !cellData.style) cellData.style = style;
        this.setCell(startRow + r, startCol + c, cellData);
      }
    }
  }

  /**
   * Iterate all cells in row-major order.
   * @returns {IterableIterator<Cell>}
   */
  *cells() {
    const keys = Array.from(this._cells.keys()).sort((a, b) => {
      const [ar, ac] = a.split(':').map(Number);
      const [br, bc] = b.split(':').map(Number);
      return ar !== br ? ar - br : ac - bc;
    });
    for (const k of keys) yield this._cells.get(k);
  }

  // ── Row / Column Properties ───────────────────────────────────────────────

  /**
   * @param {number} rowIndex
   * @param {import('../interfaces').IRowProperties} props
   */
  setRowProperties(rowIndex, props) {
    this.rowProps.set(rowIndex, { ...(this.rowProps.get(rowIndex) || {}), ...props });
  }

  /**
   * @param {number} colIndex
   * @param {import('../interfaces').IColumnProperties} props
   */
  setColProperties(colIndex, props) {
    this.colProps.set(colIndex, { ...(this.colProps.get(colIndex) || {}), ...props });
  }

  // ── Merges ────────────────────────────────────────────────────────────────

  /**
   * Merge a range. Throws if overlapping with existing merge.
   * @param {string|import('../interfaces').ICellRange} rangeOrRef
   */
  mergeCells(rangeOrRef) {
    const range = typeof rangeOrRef === 'string' ? decodeRange(rangeOrRef) : rangeOrRef;
    for (const existing of this.merges) {
      if (rangesOverlap(range, existing)) {
        throw new Error(`Merge range ${encodeRange(range.startRow, range.startCol, range.endRow, range.endCol)} overlaps with existing merge.`);
      }
    }
    this.merges.push(range);
  }

  /**
   * Remove a merge by its top-left cell.
   * @param {number} row
   * @param {number} col
   */
  unmerge(row, col) {
    this.merges = this.merges.filter(m => !(m.startRow === row && m.startCol === col));
  }

  // ── Conditional Formatting ────────────────────────────────────────────────

  /**
   * @param {import('../interfaces').IConditionalFormatting} cf
   */
  addConditionalFormatting(cf) {
    this.conditionalFormattings.push(cf);
  }

  // ── Data Validation ───────────────────────────────────────────────────────

  /**
   * @param {import('../interfaces').IDataValidation} validation
   */
  addDataValidation(validation) {
    this.dataValidations.push(validation);
  }

  // ── Auto Filter ───────────────────────────────────────────────────────────

  /**
   * @param {string|import('../interfaces').ICellRange} rangeOrRef
   * @param {Array<{colIndex:number, filters:string[]}>} [columnFilters]
   */
  setAutoFilter(rangeOrRef, columnFilters) {
    const range = typeof rangeOrRef === 'string' ? decodeRange(rangeOrRef) : rangeOrRef;
    this.autoFilter = { range, columnFilters: columnFilters || [] };
  }

  removeAutoFilter() {
    this.autoFilter = null;
  }

  // ── Tables ────────────────────────────────────────────────────────────────

  /**
   * @param {import('../interfaces').ITable} table
   */
  addTable(table) {
    this.tables.push(table);
  }

  // ── Charts ────────────────────────────────────────────────────────────────

  /**
   * @param {import('../interfaces').IChartConfig} chart
   */
  addChart(chart) {
    this.charts.push(chart);
  }

  // ── Images ────────────────────────────────────────────────────────────────

  /**
   * @param {import('../interfaces').IImage} image
   */
  addImage(image) {
    this.images.push(image);
  }

  // ── Sparklines ────────────────────────────────────────────────────────────

  /**
   * @param {import('../interfaces').ISparkline} sparkline
   */
  addSparkline(sparkline) {
    this.sparklines.push(sparkline);
  }

  // ── Groupings ─────────────────────────────────────────────────────────────

  /**
   * @param {import('../interfaces').ISheetGrouping} grouping
   */
  addGrouping(grouping) {
    this.groupings.push(grouping);
  }

  // ── Page Breaks ───────────────────────────────────────────────────────────

  /**
   * @param {number} rowIndex
   */
  addRowPageBreak(rowIndex) {
    this.rowPageBreaks.push(rowIndex);
  }

  /**
   * @param {number} colIndex
   */
  addColPageBreak(colIndex) {
    this.colPageBreaks.push(colIndex);
  }

  // ── Freeze Panes ─────────────────────────────────────────────────────────

  /**
   * @param {import('../interfaces').IFreezePanes} config
   */
  setFreezePanes(config) {
    this.sheetView = { ...this.sheetView, freezePanes: config };
  }

  // ── Protection ────────────────────────────────────────────────────────────

  /**
   * @param {import('../interfaces').IProtection} protection
   */
  setProtection(protection) {
    this.protection = protection;
  }

  // ── Internal helpers ──────────────────────────────────────────────────────

  _trackBounds(row, col) {
    if (row > this._maxRow) this._maxRow = row;
    if (col > this._maxCol) this._maxCol = col;
  }

  _resolveRef(rowOrRef, col) {
    if (typeof rowOrRef === 'string') {
      const d = decodeCell(rowOrRef);
      return { row: d.row, c: d.col };
    }
    return { row: rowOrRef, c: col };
  }

  get dimensions() {
    if (this._maxRow < 0) return null;
    return { startRow: 0, startCol: 0, endRow: this._maxRow, endCol: this._maxCol };
  }

  get cellCount() { return this._cells.size; }
}

module.exports = { Worksheet };
