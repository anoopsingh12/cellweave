/**
 * CellWeave – Public Entry Point
 * @module cellweave
 */

'use strict';

const { CellWeave, CellBuilder, WorksheetBuilder, Workbook, Worksheet } = require('./src/CellWeave');
const { StyleRegistry }       = require('./src/core/StyleRegistry');
const { SharedStringTable }   = require('./src/core/SharedStringTable');
const { XlsxWriter }          = require('./src/writers/XlsxWriter');
const addressUtils             = require('./src/utils/addressUtils');

module.exports = {
  // ── Primary API ───────────────────────────────────────────────────────────
  CellWeave,

  // ── Builders ─────────────────────────────────────────────────────────────
  CellBuilder,
  WorksheetBuilder,

  // ── Models (advanced usage) ───────────────────────────────────────────────
  Workbook,
  Worksheet,

  // ── Core registries (advanced usage) ────────────────────────────────────
  StyleRegistry,
  SharedStringTable,

  // ── Low-level writer (advanced usage) ────────────────────────────────────
  XlsxWriter,

  // ── Utility functions ─────────────────────────────────────────────────────
  utils: addressUtils,

  // ── Convenience re-exports ────────────────────────────────────────────────
  /** Quick factory: new CellWeave() */
  createWorkbook: (props) => new CellWeave(props),

  /** Quick cell builder */
  cell: () => new CellBuilder(),

  /** Convert 2D array directly to xlsx Buffer */
  fromArray: CellWeave.fromArray,

  /** Convert object array directly to xlsx Buffer */
  fromObjects: CellWeave.fromObjects,
};
