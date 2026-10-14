/**
 * @fileoverview CellWeave – CellWeave Main API (Facade Pattern / DIP)
 * Single entry point: creates workbooks, wraps them with builders,
 * and exposes a clean top-level API.
 *
 * SOLID Principles Applied:
 *  - SRP: This class is only the public API facade.
 *  - OCP: New features added via builder extension, not changing this file.
 *  - LSP: WorksheetBuilder & CellBuilder are substitutable extensions.
 *  - ISP: Users only depend on what they need (fluent API surface).
 *  - DIP: Depends on Workbook/Worksheet abstractions, not concrete writers.
 */

'use strict';

const { Workbook }          = require('./models/Workbook');
const { WorksheetBuilder }  = require('./builders/WorksheetBuilder');
const { CellBuilder }       = require('./builders/CellBuilder');
const { XlsxWriter }        = require('./writers/XlsxWriter');
const { writeFileSync }     = require('fs');

/**
 * @class CellWeave
 * The primary API for creating Excel workbooks.
 *
 * @example
 * const cw = new CellWeave({ title: 'Sales Report' });
 * const sheet = cw.addSheet('Q1');
 * sheet.writeHeaders('A1', ['Name', 'Revenue', 'Growth']);
 * sheet.write('A2', 'APAC').write('B2', 1234567).write('C2', 0.12);
 * cw.saveAs('./report.xlsx');
 */
class CellWeave {
  /**
   * @param {import('./interfaces').IWorkbookProperties} [properties]
   */
  constructor(properties = {}) {
    this._wb = new Workbook(properties);
    /** @type {WorksheetBuilder[]} */
    this._sheetBuilders = [];
  }

  // ── Sheet Management ──────────────────────────────────────────────────────

  /**
   * Add a new worksheet and return a WorksheetBuilder for it.
   * @param {string} [name='Sheet1']
   * @param {number} [insertAt]
   * @returns {WorksheetBuilder}
   */
  addSheet(name = `Sheet${this._wb.sheetCount + 1}`, insertAt) {
    const ws = this._wb.addSheet(name, insertAt);
    const builder = new WorksheetBuilder(ws);
    this._sheetBuilders.push(builder);
    return builder;
  }

  /**
   * Get a WorksheetBuilder by sheet name or index.
   * @param {string|number} nameOrIndex
   * @returns {WorksheetBuilder}
   */
  getSheet(nameOrIndex) {
    const ws = this._wb.getSheet(nameOrIndex);
    return new WorksheetBuilder(ws);
  }

  /**
   * Remove a sheet by name or index.
   * @param {string|number} nameOrIndex
   */
  removeSheet(nameOrIndex) {
    this._wb.removeSheet(nameOrIndex);
    return this;
  }

  /**
   * Copy a sheet.
   * @param {string|number} nameOrIndex
   * @param {string} newName
   * @returns {WorksheetBuilder}
   */
  copySheet(nameOrIndex, newName) {
    const ws = this._wb.copySheet(nameOrIndex, newName);
    return new WorksheetBuilder(ws);
  }

  /**
   * Move a sheet to a target position.
   * @param {string|number} nameOrIndex
   * @param {number} targetIndex
   */
  moveSheet(nameOrIndex, targetIndex) {
    this._wb.moveSheet(nameOrIndex, targetIndex);
    return this;
  }

  /**
   * Rename a sheet.
   * @param {string} oldName
   * @param {string} newName
   */
  renameSheet(oldName, newName) {
    this._wb.renameSheet(oldName, newName);
    return this;
  }

  /** @param {string|number} nameOrIndex */
  setActiveSheet(nameOrIndex) {
    this._wb.activeSheet = nameOrIndex;
    return this;
  }

  get sheetCount() { return this._wb.sheetCount; }
  get sheetNames() { return this._wb.sheets.map(s => s.name); }

  // ── Defined Names ─────────────────────────────────────────────────────────

  /**
   * @param {string} name
   * @param {string} refers  - e.g. "Sheet1!$A$1:$C$10"
   * @param {string} [scope] - sheet name for local scope
   * @param {string} [comment]
   */
  defineRange(name, refers, scope, comment) {
    this._wb.addDefinedName({ name, refers, scope, comment });
    return this;
  }

  // ── Workbook Properties ───────────────────────────────────────────────────

  /**
   * @param {import('./interfaces').IWorkbookProperties} props
   */
  setProperties(props) {
    Object.assign(this._wb.properties, props);
    return this;
  }

  /**
   * Enable 1904 date system (Mac legacy).
   */
  use1904DateSystem() {
    this._wb.properties.date1904 = true;
    return this;
  }

  // ── Output ────────────────────────────────────────────────────────────────

  /**
   * Serialise to a Uint8Array (.xlsx bytes).
   * @returns {Uint8Array}
   */
  toBuffer() {
    if (this._wb.sheetCount === 0) {
      this.addSheet('Sheet1');
    }
    const writer = new XlsxWriter(this._wb);
    return writer.build();
  }

  /**
   * Write the workbook to a file.
   * @param {string} filePath
   */
  saveAs(filePath) {
    const buf = this.toBuffer();
    writeFileSync(filePath, buf);
    return this;
  }

  /**
   * Return a Base64-encoded string of the .xlsx file.
   * @returns {string}
   */
  toBase64() {
    return Buffer.from(this.toBuffer()).toString('base64');
  }

  /**
   * Return a Blob (browser-only).
   * @returns {Blob}
   */
  toBlob() {
    const buf = this.toBuffer();
    return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  // ── Utility: cell builder factory ─────────────────────────────────────────

  /**
   * Create a new CellBuilder.
   * @returns {CellBuilder}
   */
  static cell() { return new CellBuilder(); }

  // ── Static helpers ────────────────────────────────────────────────────────

  /**
   * Quickly convert a 2D array to an xlsx Buffer.
   * @param {Array<Array<*>>} data
   * @param {Object} [options]
   * @param {string}  [options.sheetName]
   * @param {string[]} [options.headers]
   * @returns {Uint8Array}
   */
  static fromArray(data, options = {}) {
    const cw = new CellWeave();
    const sheet = cw.addSheet(options.sheetName || 'Sheet1');
    let startRow = 0;
    if (options.headers && options.headers.length) {
      sheet.writeHeaders(0, 0, options.headers);
      startRow = 1;
    }
    sheet.writeArray(startRow, 0, data);
    return cw.toBuffer();
  }

  /**
   * Quickly convert an array of objects to an xlsx Buffer.
   * Uses object keys as headers.
   * @param {Object[]} records
   * @param {Object}  [options]
   * @param {string}  [options.sheetName]
   * @param {string[]} [options.columns] - explicit column order
   * @returns {Uint8Array}
   */
  static fromObjects(records, options = {}) {
    if (!records || records.length === 0) return new CellWeave().toBuffer();
    const keys = options.columns || Object.keys(records[0]);
    const data = records.map(row => keys.map(k => row[k]));
    return CellWeave.fromArray(data, { ...options, headers: keys });
  }
}

module.exports = { CellWeave, CellBuilder, WorksheetBuilder, Workbook, Worksheet: require('./models/Worksheet').Worksheet };
