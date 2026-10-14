/**
 * @fileoverview CellWeave – StyleRegistry (Single Responsibility)
 * Manages all styles, fonts, fills, borders and number formats.
 * Deduplicates and assigns integer IDs for xlsx output.
 */

'use strict';

const { hexToArgb } = require('../utils/addressUtils');

/**
 * Canonical default font for new workbooks.
 */
const DEFAULT_FONT = {
  name: 'Calibri',
  size: 11,
  color: '000000',
  bold: false,
  italic: false,
  underline: 'none',
  strikethrough: false,
  vertAlign: 'baseline',
  family: 'Swiss',
  scheme: 'minor',
};

/**
 * Built-in number format IDs (0-49 are reserved by OOXML).
 */
const BUILTIN_FORMATS = new Map([
  [0,  'General'],
  [1,  '0'],
  [2,  '0.00'],
  [3,  '#,##0'],
  [4,  '#,##0.00'],
  [9,  '0%'],
  [10, '0.00%'],
  [11, '0.00E+00'],
  [12, '# ?/?'],
  [13, '# ??/??'],
  [14, 'mm-dd-yy'],
  [15, 'd-mmm-yy'],
  [16, 'd-mmm'],
  [17, 'mmm-yy'],
  [18, 'h:mm AM/PM'],
  [19, 'h:mm:ss AM/PM'],
  [20, 'h:mm'],
  [21, 'h:mm:ss'],
  [22, 'm/d/yy h:mm'],
  [37, '#,##0 ;(#,##0)'],
  [38, '#,##0 ;[Red](#,##0)'],
  [39, '#,##0.00;(#,##0.00)'],
  [40, '#,##0.00;[Red](#,##0.00)'],
  [45, 'mm:ss'],
  [46, '[h]:mm:ss'],
  [47, 'mmss.0'],
  [48, '##0.0E+0'],
  [49, '@'],
]);

/**
 * @class StyleRegistry
 * Implements a registry pattern to deduplicate and index style components.
 */
class StyleRegistry {
  constructor() {
    /** @private @type {Map<string, number>} font-key → index */
    this._fonts = new Map();
    /** @private @type {Array<Object>} ordered fonts */
    this._fontList = [];

    /** @private @type {Map<string, number>} */
    this._fills = new Map();
    /** @private @type {Array<Object>} */
    this._fillList = [];

    /** @private @type {Map<string, number>} */
    this._borders = new Map();
    /** @private @type {Array<Object>} */
    this._borderList = [];

    /** @private @type {Map<string, number>} format code → id */
    this._numFmts = new Map();
    /** @private @type {Map<number, string>} id → format code */
    this._numFmtById = new Map();
    /** @private @type {number} custom format IDs start at 164 */
    this._nextNumFmtId = 164;

    /** @private @type {Map<string, number>} xf-key → xfIndex */
    this._xfs = new Map();
    /** @private @type {Array<Object>} cell format records */
    this._xfList = [];

    // Seed required default entries (OOXML requires at least 1 font/fill/border)
    this._seedDefaults();

    // Register builtin formats
    for (const [id, code] of BUILTIN_FORMATS) {
      this._numFmts.set(code, id);
      this._numFmtById.set(id, code);
    }
  }

  _seedDefaults() {
    // Default font (index 0)
    this._registerFont(DEFAULT_FONT);

    // OOXML requires first two fills: none and gray125
    this._registerFill({ patternType: 'none' });
    this._registerFill({ patternType: 'gray125' });

    // Default border (no borders, index 0)
    this._registerBorder({});
  }

  // ── Font ──────────────────────────────────────────────────────────────────

  /**
   * @param {import('../interfaces').IFont} font
   * @returns {number} font index
   */
  registerFont(font) {
    return this._registerFont({ ...DEFAULT_FONT, ...font });
  }

  _registerFont(font) {
    const key = JSON.stringify(font);
    if (this._fonts.has(key)) return this._fonts.get(key);
    const idx = this._fontList.length;
    this._fonts.set(key, idx);
    this._fontList.push(font);
    return idx;
  }

  getFonts() { return this._fontList; }

  // ── Fill ──────────────────────────────────────────────────────────────────

  /**
   * @param {import('../interfaces').IFill} fill
   * @returns {number} fill index
   */
  registerFill(fill) {
    return this._registerFill(fill);
  }

  _registerFill(fill) {
    const key = JSON.stringify(fill);
    if (this._fills.has(key)) return this._fills.get(key);
    const idx = this._fillList.length;
    this._fills.set(key, idx);
    this._fillList.push(fill);
    return idx;
  }

  getFills() { return this._fillList; }

  // ── Border ────────────────────────────────────────────────────────────────

  /**
   * @param {import('../interfaces').IBorder} border
   * @returns {number} border index
   */
  registerBorder(border) {
    return this._registerBorder(border);
  }

  _registerBorder(border) {
    const key = JSON.stringify(border);
    if (this._borders.has(key)) return this._borders.get(key);
    const idx = this._borderList.length;
    this._borders.set(key, idx);
    this._borderList.push(border);
    return idx;
  }

  getBorders() { return this._borderList; }

  // ── Number Formats ────────────────────────────────────────────────────────

  /**
   * @param {string} formatCode  - Excel format string e.g. "#,##0.00"
   * @returns {number} numFmtId
   */
  registerNumberFormat(formatCode) {
    if (this._numFmts.has(formatCode)) return this._numFmts.get(formatCode);
    const id = this._nextNumFmtId++;
    this._numFmts.set(formatCode, id);
    this._numFmtById.set(id, formatCode);
    return id;
  }

  /**
   * Resolve a format type shorthand to a format code.
   * @param {import('../interfaces').INumberFormat} fmt
   * @returns {number} numFmtId
   */
  resolveNumberFormat(fmt) {
    if (!fmt) return 0;
    if (fmt.formatCode) return this.registerNumberFormat(fmt.formatCode);
    const code = StyleRegistry._buildFormatCode(fmt);
    return code ? this.registerNumberFormat(code) : 0;
  }

  /**
   * @param {import('../interfaces').INumberFormat} fmt
   * @returns {string}
   */
  static _buildFormatCode(fmt) {
    const { type, decimalPlaces = 2, useThousandSeparator, currencySymbol = '$' } = fmt;
    switch (type) {
      case 'number':
        return `${useThousandSeparator ? '#,##' : ''}0${decimalPlaces > 0 ? '.' + '0'.repeat(decimalPlaces) : ''}`;
      case 'currency':
        return `"${currencySymbol}"#,##0${decimalPlaces > 0 ? '.' + '0'.repeat(decimalPlaces) : ''}`;
      case 'accounting':
        return `_("${currencySymbol}"* #,##0${decimalPlaces > 0 ? '.' + '0'.repeat(decimalPlaces) : ''}_);_("${currencySymbol}"* (#,##0${decimalPlaces > 0 ? '.' + '0'.repeat(decimalPlaces) : ''});_("${currencySymbol}"* "-"??_);_(@_)`;
      case 'percentage':
        return `0${decimalPlaces > 0 ? '.' + '0'.repeat(decimalPlaces) : ''}%`;
      case 'scientific':
        return `0${decimalPlaces > 0 ? '.' + '0'.repeat(decimalPlaces) : ''}E+00`;
      case 'fraction':
        return '# ?/?';
      case 'date':
        return 'yyyy-mm-dd';
      case 'time':
        return 'hh:mm:ss';
      case 'datetime':
        return 'yyyy-mm-dd hh:mm:ss';
      case 'text':
        return '@';
      default:
        return '';
    }
  }

  getNumFmts() {
    const custom = [];
    for (const [id, code] of this._numFmtById) {
      if (id >= 164) custom.push({ id, code });
    }
    return custom;
  }

  // ── Cell XF (Cell Format) ─────────────────────────────────────────────────

  /**
   * Register a complete cell style and return its xfIndex.
   * @param {import('../interfaces').ICellStyle} style
   * @returns {number} xfIndex (0 = default)
   */
  registerCellStyle(style) {
    if (!style) return 0;
    const fontIdx    = style.font   ? this.registerFont(style.font)     : 0;
    const fillIdx    = style.fill   ? this.registerFill(style.fill)     : 0;
    const borderIdx  = style.border ? this.registerBorder(style.border) : 0;
    const numFmtId   = style.numberFormat ? this.resolveNumberFormat(style.numberFormat) : 0;
    const alignment  = style.alignment || null;
    const locked     = style.locked !== undefined ? style.locked : true;
    const hidden     = style.hidden !== undefined ? style.hidden : false;

    const xf = { fontIdx, fillIdx, borderIdx, numFmtId, alignment, locked, hidden };
    const key = JSON.stringify(xf);
    if (this._xfs.has(key)) return this._xfs.get(key);

    // Index 0 is seeded as default — if first call returns index 0 we're fine.
    if (this._xfList.length === 0) {
      this._xfList.push({ fontIdx: 0, fillIdx: 0, borderIdx: 0, numFmtId: 0, alignment: null, locked: true, hidden: false });
      this._xfs.set(JSON.stringify(this._xfList[0]), 0);
    }

    const idx = this._xfList.length;
    this._xfs.set(key, idx);
    this._xfList.push(xf);
    return idx;
  }

  getCellFormats() { return this._xfList; }
}

module.exports = { StyleRegistry, DEFAULT_FONT, BUILTIN_FORMATS };
