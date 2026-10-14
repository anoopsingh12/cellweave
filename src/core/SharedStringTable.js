/**
 * @fileoverview CellWeave – SharedStringTable (Single Responsibility)
 * Manages Excel shared strings for efficient string deduplication.
 */

'use strict';

const { escapeXml } = require('../utils/addressUtils');

/**
 * @class SharedStringTable
 * Deduplicates cell string values and assigns integer indexes for xlsx output.
 * Rich-text (multi-run) strings are also supported.
 */
class SharedStringTable {
  constructor() {
    /** @private @type {Map<string, number>} */
    this._map = new Map();
    /** @private @type {Array<string|Array>} */
    this._strings = [];
    /** @private @type {number} */
    this._totalCount = 0;
  }

  /**
   * Register a plain string. Returns its shared-string index.
   * @param {string} str
   * @returns {number}
   */
  addString(str) {
    this._totalCount++;
    const key = `plain:${str}`;
    if (this._map.has(key)) return this._map.get(key);
    const idx = this._strings.length;
    this._map.set(key, idx);
    this._strings.push({ type: 'plain', value: str });
    return idx;
  }

  /**
   * Register a rich-text string (array of runs).
   * @param {Array<{text:string, font?:import('../interfaces').IFont}>} runs
   * @returns {number}
   */
  addRichText(runs) {
    this._totalCount++;
    const key = `rich:${JSON.stringify(runs)}`;
    if (this._map.has(key)) return this._map.get(key);
    const idx = this._strings.length;
    this._map.set(key, idx);
    this._strings.push({ type: 'rich', runs });
    return idx;
  }

  /** @returns {number} */
  get uniqueCount() { return this._strings.length; }

  /** @returns {number} */
  get totalCount() { return this._totalCount; }

  /** @returns {Array<{type:string, value?:string, runs?:Array}>} */
  getStrings() { return this._strings; }
}

module.exports = { SharedStringTable };
