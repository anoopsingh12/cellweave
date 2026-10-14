/**
 * @fileoverview CellWeave – Cell Model (Single Responsibility)
 * Represents a single cell with value, type, formula, style, comment, hyperlink, etc.
 */

'use strict';

const { encodeCell, dateToSerial } = require('../utils/addressUtils');

/**
 * @class Cell
 * Immutable-friendly data model for a single spreadsheet cell.
 */
class Cell {
  /**
   * @param {number} row  - 0-based row
   * @param {number} col  - 0-based col
   * @param {import('../interfaces').ICellData} [data]
   */
  constructor(row, col, data = {}) {
    this.row = row;
    this.col = col;
    this._data = { ...data };
  }

  // ── Value & Type ──────────────────────────────────────────────────────────

  get value() { return this._data.value; }
  get type()  { return this._data.type;  }
  get formula() { return this._data.formula; }
  get style()   { return this._data.style; }
  get comment() { return this._data.comment; }
  get hyperlink() { return this._data.hyperlink; }
  get isArrayFormula() { return this._data.isArrayFormula; }
  get arrayFormulaRange() { return this._data.arrayFormulaRange; }
  get richText() { return this._data.richText; }

  /**
   * Resolve the effective cell type if not explicitly provided.
   * @returns {import('../interfaces').CellType}
   */
  resolvedType() {
    if (this._data.type) return this._data.type;
    if (this._data.formula) return 'str';
    const v = this._data.value;
    if (v === null || v === undefined) return 'z';
    if (typeof v === 'boolean') return 'b';
    if (typeof v === 'number') return 'n';
    if (v instanceof Date) return 'n'; // dates stored as numbers
    if (typeof v === 'string') return 's'; // shared string
    return 'str';
  }

  /**
   * Get the numeric value for numbers/booleans/dates.
   * @param {boolean} [date1904=false]
   * @returns {number|null}
   */
  numericValue(date1904 = false) {
    const v = this._data.value;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (v instanceof Date) return dateToSerial(v, date1904);
    return null;
  }

  /**
   * Merge another data object on top of this cell (for style updates etc.)
   * @param {import('../interfaces').ICellData} patch
   * @returns {Cell}
   */
  merge(patch) {
    return new Cell(this.row, this.col, { ...this._data, ...patch });
  }

  /**
   * @returns {string} A1 address
   */
  get address() {
    return encodeCell(this.row, this.col);
  }

  toJSON() {
    return { row: this.row, col: this.col, ...this._data };
  }
}

module.exports = { Cell };
