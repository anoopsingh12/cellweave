/**
 * @fileoverview CellWeave – Workbook Model (Single Responsibility / Open-Closed)
 * The top-level container for all worksheets, named ranges, and workbook properties.
 */

'use strict';

const { Worksheet } = require('./Worksheet');
const { StyleRegistry } = require('../core/StyleRegistry');
const { SharedStringTable } = require('../core/SharedStringTable');
const { sanitizeSheetName } = require('../utils/addressUtils');

/**
 * @class Workbook
 * Represents a complete Excel workbook (.xlsx).
 * Follows Open-Closed: extend workbook features via plugins, not modification.
 */
class Workbook {
  /**
   * @param {import('../interfaces').IWorkbookProperties} [properties]
   */
  constructor(properties = {}) {
    /** @type {import('../interfaces').IWorkbookProperties} */
    this.properties = {
      title: '',
      subject: '',
      author: 'Anoop Singh',
      company: '',
      category: '',
      keywords: '',
      description: '',
      language: 'en-US',
      created: new Date(),
      modified: new Date(),
      lastModifiedBy: 'CellWeave',
      date1904: false,
      ...properties,
    };

    /** @type {Worksheet[]} */
    this._sheets = [];

    /** @type {Map<string, Worksheet>} */
    this._sheetsByName = new Map();

    /** @type {StyleRegistry} */
    this.styles = new StyleRegistry();

    /** @type {SharedStringTable} */
    this.sharedStrings = new SharedStringTable();

    /** @type {Array<import('../interfaces').IDefinedName>} */
    this.definedNames = [];

    /** @type {Map<string, any>} plugin data storage */
    this._pluginData = new Map();

    /** @type {number} */
    this._activeSheetIndex = 0;
  }

  // ── Sheet Management ──────────────────────────────────────────────────────

  /**
   * Add a new worksheet.
   * @param {string} name
   * @param {number} [insertAt] - 0-based index; defaults to append
   * @returns {Worksheet}
   */
  addSheet(name, insertAt) {
    const safeName = sanitizeSheetName(name);
    if (this._sheetsByName.has(safeName)) {
      throw new Error(`Sheet "${safeName}" already exists`);
    }
    const sheetId = this._sheets.length + 1;
    const ws = new Worksheet(safeName, sheetId);

    if (insertAt !== undefined && insertAt >= 0 && insertAt <= this._sheets.length) {
      this._sheets.splice(insertAt, 0, ws);
      // Rebuild map
      this._sheetsByName.clear();
      this._sheets.forEach(s => this._sheetsByName.set(s.name, s));
    } else {
      this._sheets.push(ws);
      this._sheetsByName.set(safeName, ws);
    }
    return ws;
  }

  /**
   * Get a sheet by name or index.
   * @param {string|number} nameOrIndex
   * @returns {Worksheet}
   */
  getSheet(nameOrIndex) {
    if (typeof nameOrIndex === 'number') {
      if (nameOrIndex < 0 || nameOrIndex >= this._sheets.length) {
        throw new RangeError(`Sheet index ${nameOrIndex} out of range`);
      }
      return this._sheets[nameOrIndex];
    }
    const ws = this._sheetsByName.get(nameOrIndex);
    if (!ws) throw new Error(`Sheet "${nameOrIndex}" not found`);
    return ws;
  }

  /**
   * Remove a sheet by name or index.
   * @param {string|number} nameOrIndex
   */
  removeSheet(nameOrIndex) {
    const ws = this.getSheet(nameOrIndex);
    this._sheets = this._sheets.filter(s => s !== ws);
    this._sheetsByName.delete(ws.name);
  }

  /**
   * Rename a sheet.
   * @param {string} oldName
   * @param {string} newName
   */
  renameSheet(oldName, newName) {
    const ws = this.getSheet(oldName);
    const safe = sanitizeSheetName(newName);
    if (this._sheetsByName.has(safe) && safe !== oldName) {
      throw new Error(`Sheet "${safe}" already exists`);
    }
    this._sheetsByName.delete(ws.name);
    ws.name = safe;
    this._sheetsByName.set(safe, ws);
  }

  /**
   * Move a sheet to a new position.
   * @param {string|number} nameOrIndex
   * @param {number} targetIndex
   */
  moveSheet(nameOrIndex, targetIndex) {
    const ws = this.getSheet(nameOrIndex);
    this._sheets = this._sheets.filter(s => s !== ws);
    this._sheets.splice(targetIndex, 0, ws);
  }

  /**
   * Copy a sheet (deep clone its cells, properties, merges etc.)
   * @param {string|number} nameOrIndex
   * @param {string} newName
   * @returns {Worksheet}
   */
  copySheet(nameOrIndex, newName) {
    const source = this.getSheet(nameOrIndex);
    const target = this.addSheet(newName);
    // Deep clone via JSON (sufficient for data; images/charts are refs)
    for (const cell of source.cells()) {
      target.setCell(cell.row, cell.col, cell.toJSON());
    }
    target.rowProps = new Map(JSON.parse(JSON.stringify([...source.rowProps])));
    target.colProps = new Map(JSON.parse(JSON.stringify([...source.colProps])));
    target.merges  = JSON.parse(JSON.stringify(source.merges));
    target.pageSetup = { ...source.pageSetup };
    target.pageMargins = { ...source.pageMargins };
    target.sheetView = JSON.parse(JSON.stringify(source.sheetView));
    return target;
  }

  get sheets() { return [...this._sheets]; }
  get sheetCount() { return this._sheets.length; }

  // ── Active Sheet ──────────────────────────────────────────────────────────

  set activeSheet(nameOrIndex) {
    const ws = this.getSheet(nameOrIndex);
    this._activeSheetIndex = this._sheets.indexOf(ws);
  }

  get activeSheet() { return this._sheets[this._activeSheetIndex]; }

  // ── Defined Names ─────────────────────────────────────────────────────────

  /**
   * @param {import('../interfaces').IDefinedName} definition
   */
  addDefinedName(definition) {
    const existing = this.definedNames.findIndex(d => d.name === definition.name && d.scope === definition.scope);
    if (existing >= 0) {
      this.definedNames[existing] = definition;
    } else {
      this.definedNames.push(definition);
    }
  }

  /**
   * @param {string} name
   * @param {string} [scope]
   */
  removeDefinedName(name, scope) {
    this.definedNames = this.definedNames.filter(d => !(d.name === name && d.scope === scope));
  }

  // ── Plugin Data Store ─────────────────────────────────────────────────────

  /**
   * Store arbitrary plugin data by key.
   * @param {string} key
   * @param {*} data
   */
  setPluginData(key, data) { this._pluginData.set(key, data); }

  /**
   * @param {string} key
   * @returns {*}
   */
  getPluginData(key) { return this._pluginData.get(key); }
}

module.exports = { Workbook };
