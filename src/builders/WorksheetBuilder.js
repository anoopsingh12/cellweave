/**
 * @fileoverview CellWeave – WorksheetBuilder (Fluent Builder Pattern)
 * Provides a high-level fluent API over the Worksheet model.
 * Follows OCP: extend capabilities without modifying Worksheet core.
 */

'use strict';

const { CellBuilder }  = require('./CellBuilder');
const { decodeCell, decodeRange } = require('../utils/addressUtils');

/**
 * @class WorksheetBuilder
 * Wraps a Worksheet model and exposes chainable configuration methods.
 */
class WorksheetBuilder {
  /**
   * @param {import('../models/Worksheet').Worksheet} ws
   */
  constructor(ws) {
    this._ws = ws;
  }

  /** Access the underlying Worksheet. */
  get sheet() { return this._ws; }

  // ── Cell Writing ──────────────────────────────────────────────────────────

  /**
   * Write a single value.
   * @param {string|number} ref  - A1 ref or row index
   * @param {number|*}      colOrValue
   * @param {*}             [value]
   */
  write(ref, colOrValue, value) {
    if (typeof ref === 'string' && arguments.length === 2) {
      const { row, col } = decodeCell(ref);
      const data = typeof colOrValue === 'object' && !(colOrValue instanceof Date)
        ? colOrValue
        : { value: colOrValue };
      this._ws.setCell(row, col, data);
    } else {
      const data = (value !== undefined && typeof value !== 'object') || value instanceof Date
        ? { value }
        : (value || {});
      this._ws.setCell(ref, colOrValue, data);
    }
    return this;
  }

  /**
   * Write a CellBuilder result or a plain ICellData object.
   * @param {string} ref - A1 ref
   * @param {CellBuilder|import('../interfaces').ICellData} builderOrData
   */
  writeCell(ref, builderOrData) {
    const { row, col } = decodeCell(ref);
    const data = (builderOrData && typeof builderOrData.build === 'function')
      ? builderOrData.build()
      : builderOrData;
    this._ws.setCell(row, col, data);
    return this;
  }

  /**
   * Write a 2D data array.
   * @param {string|number} startRef
   * @param {number|Array}  colOrData
   * @param {Array}         [data]
   * @param {import('../interfaces').ICellStyle} [style]
   */
  writeArray(startRef, colOrData, data, style) {
    if (typeof startRef === 'string') {
      const { row, col } = decodeCell(startRef);
      this._ws.writeArray(row, col, colOrData, style);
    } else {
      this._ws.writeArray(startRef, colOrData, data, style);
    }
    return this;
  }

  /**
   * Write headers as a row with bold style.
   * @param {string|number} refOrRow  - A1 reference or 0-based row index
   * @param {number|string[]} colOrHeaders - column index (when refOrRow is a number) or headers array
   * @param {string[]|import('../interfaces').ICellStyle} [headersOrStyle]
   * @param {import('../interfaces').ICellStyle} [style]
   */
  writeHeaders(refOrRow, colOrHeaders, headersOrStyle, style) {
    let row, col, headers, styleObj;
    if (typeof refOrRow === 'string') {
      // writeHeaders('A1', ['H1','H2'], style?)
      const d = decodeCell(refOrRow);
      row = d.row; col = d.col;
      headers = colOrHeaders;
      styleObj = headersOrStyle;
    } else if (typeof colOrHeaders === 'number') {
      // writeHeaders(rowIdx, colIdx, ['H1','H2'], style?)
      row = refOrRow; col = colOrHeaders;
      headers = headersOrStyle;
      styleObj = style;
    } else {
      // writeHeaders(rowIdx, ['H1','H2'], style?)
      row = refOrRow; col = 0;
      headers = colOrHeaders;
      styleObj = headersOrStyle;
    }
    headers.forEach((h, i) => {
      this._ws.setCell(row, col + i, {
        value: h,
        style: styleObj || {
          font: { bold: true, color: 'FFFFFF' },
          fill: { patternType: 'solid', fgColor: '4472C4' },
          alignment: { horizontal: 'center', vertical: 'center' },
        },
      });
    });
    return this;
  }

  // ── Batch formula row ──────────────────────────────────────────────────────

  /**
   * Write formulas across a row.
   * @param {number} row
   * @param {number} startCol
   * @param {string[]} formulas
   */
  writeFormulas(row, startCol, formulas) {
    formulas.forEach((f, i) => {
      this._ws.setCell(row, startCol + i, { formula: f });
    });
    return this;
  }

  // ── Row / Column Config ───────────────────────────────────────────────────

  /** @param {number} rowIndex @param {import('../interfaces').IRowProperties} props */
  rowHeight(rowIndex, height)    { this._ws.setRowProperties(rowIndex, { height }); return this; }
  hideRow(rowIndex)              { this._ws.setRowProperties(rowIndex, { hidden: true }); return this; }
  showRow(rowIndex)              { this._ws.setRowProperties(rowIndex, { hidden: false }); return this; }
  rowOutlineLevel(idx, level)   { this._ws.setRowProperties(idx, { outlineLevel: level }); return this; }

  /** @param {number} colIndex @param {number} width */
  colWidth(colIndex, width)      { this._ws.setColProperties(colIndex, { width }); return this; }
  hideCol(colIndex)              { this._ws.setColProperties(colIndex, { hidden: true }); return this; }
  showCol(colIndex)              { this._ws.setColProperties(colIndex, { hidden: false }); return this; }
  colBestFit(colIndex)          { this._ws.setColProperties(colIndex, { bestFit: true }); return this; }
  colOutlineLevel(idx, level)   { this._ws.setColProperties(idx, { outlineLevel: level }); return this; }

  // ── Merges ────────────────────────────────────────────────────────────────

  /** @param {string} rangeRef - e.g. "A1:C3" */
  merge(rangeRef) { this._ws.mergeCells(rangeRef); return this; }
  unmerge(rangeRef) {
    const r = decodeRange(rangeRef);
    this._ws.unmerge(r.startRow, r.startCol);
    return this;
  }

  // ── Freeze Panes ──────────────────────────────────────────────────────────

  /** @param {number} rows @param {number} [cols] */
  freezeRows(rows, cols = 0) { this._ws.setFreezePanes({ rows, cols }); return this; }
  freezeCols(cols, rows = 0) { this._ws.setFreezePanes({ cols, rows }); return this; }
  freezeAt(rows, cols)       { this._ws.setFreezePanes({ rows, cols }); return this; }

  // ── Auto Filter ───────────────────────────────────────────────────────────

  /** @param {string} rangeRef */
  autoFilter(rangeRef) { this._ws.setAutoFilter(rangeRef); return this; }
  removeAutoFilter()   { this._ws.removeAutoFilter(); return this; }

  // ── Conditional Formatting ────────────────────────────────────────────────

  /** @param {import('../interfaces').IConditionalFormatting} cf */
  conditionalFormat(cf) { this._ws.addConditionalFormatting(cf); return this; }

  /**
   * Shorthand: color scale on a range.
   * @param {string} rangeRef
   * @param {string[]} colors - min, mid (optional), max
   */
  colorScale(rangeRef, colors = ['#F8696B', '#FFEB84', '#63BE7B']) {
    const range = decodeRange(rangeRef);
    const cfvo = colors.length === 2
      ? [{ type: 'min' }, { type: 'max' }]
      : [{ type: 'min' }, { type: 'percentile', val: 50 }, { type: 'max' }];
    this._ws.addConditionalFormatting({
      range,
      rules: [{ type: 'colorScale', priority: 1, colorScale: { cfvo, color: colors } }],
    });
    return this;
  }

  /**
   * Shorthand: highlight cells > value.
   * @param {string} rangeRef
   * @param {number} value
   * @param {import('../interfaces').ICellStyle} style
   */
  highlightGreaterThan(rangeRef, value, style = { fill: { patternType: 'solid', fgColor: 'FFEB84' } }) {
    const range = decodeRange(rangeRef);
    this._ws.addConditionalFormatting({
      range,
      rules: [{ type: 'cellIs', operator: 'greaterThan', formula: [String(value)], style, priority: 1 }],
    });
    return this;
  }

  /**
   * Data bar conditional formatting.
   * @param {string} rangeRef
   * @param {string} [color='#638EC6']
   */
  dataBar(rangeRef, color = '#638EC6') {
    const range = decodeRange(rangeRef);
    this._ws.addConditionalFormatting({
      range,
      rules: [{
        type: 'dataBar', priority: 1,
        dataBar: { cfvo: [{ type: 'min' }, { type: 'max' }], color },
      }],
    });
    return this;
  }

  // ── Data Validation ───────────────────────────────────────────────────────

  /**
   * Dropdown list validation.
   * @param {string} rangeRef
   * @param {string[]} list
   */
  dropdownList(rangeRef, list) {
    const range = decodeRange(rangeRef);
    this._ws.addDataValidation({
      range,
      type: 'list',
      formula1: `"${list.join(',')}"`,
      showDropDown: false,
      allowBlank: true,
    });
    return this;
  }

  /**
   * Number range validation.
   * @param {string} rangeRef
   * @param {number} min
   * @param {number} max
   */
  validateNumberRange(rangeRef, min, max) {
    const range = decodeRange(rangeRef);
    this._ws.addDataValidation({
      range,
      type: 'decimal',
      operator: 'between',
      formula1: String(min),
      formula2: String(max),
      showErrorAlert: true,
      errorStyle: 'stop',
      errorTitle: 'Invalid Input',
      error: `Please enter a number between ${min} and ${max}`,
    });
    return this;
  }

  /** @param {import('../interfaces').IDataValidation} dv */
  dataValidation(dv) { this._ws.addDataValidation(dv); return this; }

  // ── Tables ────────────────────────────────────────────────────────────────

  /** @param {import('../interfaces').ITable} table */
  addTable(table) { this._ws.addTable(table); return this; }

  // ── Charts ────────────────────────────────────────────────────────────────

  /** @param {import('../interfaces').IChartConfig} chart */
  addChart(chart) { this._ws.addChart(chart); return this; }

  // ── Images ────────────────────────────────────────────────────────────────

  /** @param {import('../interfaces').IImage} image */
  addImage(image) { this._ws.addImage(image); return this; }

  // ── Sparklines ────────────────────────────────────────────────────────────

  /** @param {import('../interfaces').ISparkline} sparkline */
  addSparkline(sparkline) { this._ws.addSparkline(sparkline); return this; }

  // ── Groupings ─────────────────────────────────────────────────────────────

  groupRows(startIndex, endIndex, collapsed = false) {
    this._ws.addGrouping({ type: 'row', startIndex, endIndex, collapsed });
    return this;
  }
  groupCols(startIndex, endIndex, collapsed = false) {
    this._ws.addGrouping({ type: 'col', startIndex, endIndex, collapsed });
    return this;
  }

  // ── Page Breaks ───────────────────────────────────────────────────────────

  pageBreakAfterRow(rowIndex) { this._ws.addRowPageBreak(rowIndex); return this; }
  pageBreakAfterCol(colIndex) { this._ws.addColPageBreak(colIndex); return this; }

  // ── Page Setup ────────────────────────────────────────────────────────────

  /** @param {import('../interfaces').IPageSetup} setup */
  pageSetup(setup) { Object.assign(this._ws.pageSetup, setup); return this; }
  landscape() { this._ws.pageSetup.orientation = 'landscape'; return this; }
  portrait()  { this._ws.pageSetup.orientation = 'portrait'; return this; }
  fitToPage(width = 1, height = 1) {
    this._ws.pageSetup.fitToPage = 1;
    this._ws.pageSetup.fitToWidth  = width;
    this._ws.pageSetup.fitToHeight = height;
    return this;
  }
  paperA4()     { this._ws.pageSetup.paperSize = 'A4'; return this; }
  paperLetter() { this._ws.pageSetup.paperSize = 'letter'; return this; }
  printScale(pct) { this._ws.pageSetup.scale = pct; return this; }

  /** @param {import('../interfaces').IPageMargins} margins */
  pageMargins(m) { Object.assign(this._ws.pageMargins, m); return this; }

  // ── Header / Footer ───────────────────────────────────────────────────────

  /** @param {import('../interfaces').IHeaderFooter} hf */
  headerFooter(hf) { Object.assign(this._ws.headerFooter, hf); return this; }

  // ── View ──────────────────────────────────────────────────────────────────

  hideGridLines()    { this._ws.sheetView.showGridLines = false; return this; }
  showGridLines()    { this._ws.sheetView.showGridLines = true;  return this; }
  zoom(pct)          { this._ws.sheetView.zoomScale = pct; return this; }
  showFormulas(v=true){ this._ws.sheetView.showFormulas = v; return this; }
  rightToLeft(v=true) { this._ws.sheetView.rightToLeft = v; return this; }

  // ── Tab ───────────────────────────────────────────────────────────────────

  /** @param {string} hex */
  tabColor(hex) { this._ws.tabColor = hex; return this; }

  // ── Protection ────────────────────────────────────────────────────────────

  /** @param {string} [password] @param {import('../interfaces').IProtection} [options] */
  protect(password, options = {}) {
    this._ws.setProtection({ sheet: true, password, ...options });
    return this;
  }
  unprotect() { this._ws.setProtection(null); return this; }
}

module.exports = { WorksheetBuilder };
