/**
 * @fileoverview CellWeave – CellBuilder (Fluent Builder Pattern)
 * Provides a chainable API for constructing cell data objects.
 * Follows SRP: only builds cell data; does not write to sheets.
 */

'use strict';

/**
 * @class CellBuilder
 * Fluent builder for constructing ICellData objects.
 * @example
 * const data = new CellBuilder()
 *   .value('Hello')
 *   .bold().fontSize(14).color('#FF0000')
 *   .wrapText()
 *   .build();
 */
class CellBuilder {
  constructor() {
    this._data = {};
  }

  // ── Value ─────────────────────────────────────────────────────────────────

  /** @param {*} val */
  value(val) { this._data.value = val; return this; }

  /** @param {string} formula - without leading = */
  formula(formula) { this._data.formula = formula; return this; }

  /**
   * Set as array formula over a range.
   * @param {import('../interfaces').ICellRange} range
   */
  arrayFormula(formula, range) {
    this._data.formula = formula;
    this._data.isArrayFormula = true;
    this._data.arrayFormulaRange = range;
    return this;
  }

  /** @param {import('../interfaces').CellType} type */
  type(type) { this._data.type = type; return this; }

  // ── Rich Text ─────────────────────────────────────────────────────────────

  /**
   * Set rich text runs.
   * @param {Array<{text:string, font?:import('../interfaces').IFont}>} runs
   */
  richText(runs) {
    this._data.richText = runs;
    this._data.type = 's';
    return this;
  }

  // ── Comment ───────────────────────────────────────────────────────────────

  /** @param {string} text */
  comment(text) { this._data.comment = text; return this; }

  // ── Hyperlink ─────────────────────────────────────────────────────────────

  /** @param {string} url */
  hyperlink(url) { this._data.hyperlink = url; return this; }

  // ── Style shortcuts ───────────────────────────────────────────────────────

  _getStyle()   { if (!this._data.style) this._data.style = {}; return this._data.style; }
  _getFont()    { const s = this._getStyle(); if (!s.font) s.font = {}; return s.font; }
  _getFill()    { const s = this._getStyle(); if (!s.fill) s.fill = {}; return s.fill; }
  _getBorder()  { const s = this._getStyle(); if (!s.border) s.border = {}; return s.border; }
  _getAlign()   { const s = this._getStyle(); if (!s.alignment) s.alignment = {}; return s.alignment; }
  _getNumFmt()  { const s = this._getStyle(); if (!s.numberFormat) s.numberFormat = {}; return s.numberFormat; }

  // ── Font ──────────────────────────────────────────────────────────────────

  bold(v = true)           { this._getFont().bold = v; return this; }
  italic(v = true)         { this._getFont().italic = v; return this; }
  underline(style = 'single') { this._getFont().underline = style; return this; }
  strikethrough(v = true)  { this._getFont().strikethrough = v; return this; }
  fontName(name)           { this._getFont().name = name; return this; }
  fontSize(size)           { this._getFont().size = size; return this; }
  /** @param {string} hex - e.g. "#FF0000" or "FF0000" */
  color(hex)               { this._getFont().color = hex.replace('#', ''); return this; }
  superscript()            { this._getFont().vertAlign = 'superscript'; return this; }
  subscript()              { this._getFont().vertAlign = 'subscript'; return this; }

  // ── Fill ──────────────────────────────────────────────────────────────────

  /** @param {string} hex */
  bgColor(hex) {
    const fill = this._getFill();
    fill.patternType = 'solid';
    fill.fgColor = hex.replace('#', '');
    return this;
  }

  /**
   * Gradient fill.
   * @param {'linear'|'path'} type
   * @param {number} degree
   * @param {Array<{position:number, color:string}>} stops
   */
  gradient(type, degree, stops) {
    const fill = this._getFill();
    fill.gradient = { type, degree, stops: stops.map(s => ({ position: s.position, color: s.color.replace('#', '') })) };
    return this;
  }

  /** @param {import('../interfaces').FillPattern} pattern */
  fillPattern(pattern, fgColor, bgColor) {
    const fill = this._getFill();
    fill.patternType = pattern;
    if (fgColor) fill.fgColor = fgColor.replace('#', '');
    if (bgColor) fill.bgColor = bgColor.replace('#', '');
    return this;
  }

  // ── Border ────────────────────────────────────────────────────────────────

  /**
   * Set all borders at once.
   * @param {import('../interfaces').BorderStyle} style
   * @param {string} [color]
   */
  allBorders(style, color) {
    const b = this._getBorder();
    const side = { style, color: color ? color.replace('#', '') : undefined };
    b.top = b.bottom = b.left = b.right = { ...side };
    return this;
  }

  /** @param {import('../interfaces').BorderStyle} style */
  borderTop(style, color)    { this._getBorder().top    = { style, color: color?.replace('#', '') }; return this; }
  borderBottom(style, color) { this._getBorder().bottom = { style, color: color?.replace('#', '') }; return this; }
  borderLeft(style, color)   { this._getBorder().left   = { style, color: color?.replace('#', '') }; return this; }
  borderRight(style, color)  { this._getBorder().right  = { style, color: color?.replace('#', '') }; return this; }

  // ── Alignment ─────────────────────────────────────────────────────────────

  /** @param {import('../interfaces').HorizontalAlignment} h */
  alignH(h)           { this._getAlign().horizontal = h; return this; }
  /** @param {import('../interfaces').VerticalAlignment} v */
  alignV(v)           { this._getAlign().vertical = v; return this; }
  wrapText(v = true)  { this._getAlign().wrapText = v; return this; }
  shrinkToFit(v=true) { this._getAlign().shrinkToFit = v; return this; }
  /** @param {number} degrees - -90 to 90 */
  rotate(degrees)     { this._getAlign().rotation = degrees; return this; }
  indent(n)           { this._getAlign().indent = n; return this; }

  // ── Number Format ─────────────────────────────────────────────────────────

  /** @param {string} code - Excel format code e.g. "#,##0.00" */
  formatCode(code)    { this._getNumFmt().formatCode = code; return this; }
  formatCurrency(symbol = '$', decimals = 2) {
    const nf = this._getNumFmt();
    nf.type = 'currency'; nf.currencySymbol = symbol; nf.decimalPlaces = decimals;
    return this;
  }
  formatDate()        { this._getNumFmt().type = 'date'; return this; }
  formatTime()        { this._getNumFmt().type = 'time'; return this; }
  formatDateTime()    { this._getNumFmt().type = 'datetime'; return this; }
  formatPercent(d=2)  { this._getNumFmt().type = 'percentage'; this._getNumFmt().decimalPlaces = d; return this; }
  formatNumber(d=2, thousands=false) {
    const nf = this._getNumFmt();
    nf.type = 'number'; nf.decimalPlaces = d; nf.useThousandSeparator = thousands;
    return this;
  }
  formatAccounting(symbol='$', d=2) {
    const nf = this._getNumFmt();
    nf.type = 'accounting'; nf.currencySymbol = symbol; nf.decimalPlaces = d;
    return this;
  }
  formatText()        { this._getNumFmt().type = 'text'; return this; }
  formatScientific(d=2) { this._getNumFmt().type = 'scientific'; this._getNumFmt().decimalPlaces = d; return this; }
  formatFraction()    { this._getNumFmt().type = 'fraction'; return this; }

  // ── Protection ────────────────────────────────────────────────────────────

  lock(v = true)   { this._getStyle().locked = v; return this; }
  hide(v = true)   { this._getStyle().hidden = v; return this; }

  // ── Style object ──────────────────────────────────────────────────────────

  /** Apply a raw style object (merged over current). */
  style(styleObj) {
    this._data.style = this._deepMerge(this._data.style || {}, styleObj);
    return this;
  }

  // ── Build ─────────────────────────────────────────────────────────────────

  /**
   * @returns {import('../interfaces').ICellData}
   */
  build() {
    return { ...this._data };
  }

  _deepMerge(target, source) {
    const result = { ...target };
    for (const key of Object.keys(source)) {
      if (source[key] && typeof source[key] === 'object' && !Array.isArray(source[key])) {
        result[key] = this._deepMerge(result[key] || {}, source[key]);
      } else {
        result[key] = source[key];
      }
    }
    return result;
  }
}

module.exports = { CellBuilder };
