/**
 * ============================================================
 *  0b-overflow-manager.js  —  Cell Overflow Manager
 *  Expense Management System (GAS)
 *
 *  Loaded after 0-shared_utils.js (alphabetically safe).
 *  Transparently handles Google Sheets' 32,767-char cell limit
 *  by spilling oversized JSON fields into a dedicated
 *  "OverflowStore" sheet, storing a pointer in the main cell.
 *
 *  Pointer format (stored in main Expenses cell):
 *    __OVF:<referenceNumber>:<fieldName>
 *  Example:
 *    __OVF:REF-20250407-001:historyJSON
 *
 *  OverflowStore sheet columns:
 *    referenceNumber | fieldName | data | updatedAt
 *
 *  Integration points (only 2 files need small edits):
 *    • 0-shared_utils.js → SharedHelper.stringifyJsonFields()
 *                          SharedHelper.parseJsonFields()
 *    • initializeApp()   → call OverflowManager.ensureSheet()
 *
 *  All other layers (UnifiedService, RPCs, frontend) are 100%
 *  unaffected — they always see fully-merged objects.
 * ============================================================
 */

// ── CONSTANTS ─────────────────────────────────────────────────
var OVERFLOW = {
  SHEET_NAME:   'OverflowStore',
  POINTER_PREFIX: '__OVF:',
  // Trigger spill at 28 000 chars — leaves ~4 700 char safety buffer
  // below the hard 32 767 limit, accounting for other non-JSON columns.
  THRESHOLD:    28000,

  // Column layout of OverflowStore (1-based for setValues)
  COL: {
    REF_NUMBER: 1,   // referenceNumber
    FIELD_NAME: 2,   // e.g. "historyJSON"
    DATA:       3,   // full JSON string
    UPDATED_AT: 4    // ISO timestamp
  },

  HEADERS: ['referenceNumber', 'fieldName', 'data', 'updatedAt']
};

// ── OVERFLOW MANAGER CLASS ────────────────────────────────────
/**
 * OverflowManager
 *
 * Injected into SharedHelper so it can be called from
 * stringifyJsonFields() (write path) and parseJsonFields() (read path).
 *
 * Constructor receives the same `db` (Spreadsheet) that SharedHelper uses.
 */
class OverflowManager {

  constructor(db) {
    this.db = db;
    this._sheet = null; // lazy-loaded
  }

  // ── Sheet bootstrap ─────────────────────────────────────────
  /**
   * Call once from initializeApp() to create the OverflowStore sheet.
   * Safe to call again — no-ops if already present.
   */
  ensureSheet() {
    let sheet = this.db.getSheetByName(OVERFLOW.SHEET_NAME);
    if (sheet) {
      Logger.log('OverflowStore sheet already exists — skipped.');
      return sheet;
    }

    sheet = this.db.insertSheet(OVERFLOW.SHEET_NAME);
    sheet.appendRow(OVERFLOW.HEADERS);

    // Style header row
    const hRange = sheet.getRange(1, 1, 1, OVERFLOW.HEADERS.length);
    hRange.setFontWeight('bold');
    hRange.setBackground('#744210');   // amber-900 — visually distinct
    hRange.setFontColor('#FFFFFF');
    hRange.setWrap(false);

    // Wide data column for readability
    sheet.setColumnWidth(OVERFLOW.COL.DATA, 600);
    sheet.setFrozenRows(1);

    Logger.log('OverflowStore sheet created.');
    return sheet;
  }

  // ── Internal: lazy sheet getter ─────────────────────────────
  _getSheet() {
    if (!this._sheet) {
      this._sheet = this.db.getSheetByName(OVERFLOW.SHEET_NAME);
      if (!this._sheet) {
        // Auto-heal: create if missing (e.g. accidentally deleted)
        Logger.log('WARNING: OverflowStore missing — auto-creating.');
        this._sheet = this.ensureSheet();
      }
    }
    return this._sheet;
  }

  // ── Pointer helpers ─────────────────────────────────────────
  /**
   * Build a pointer string for a given ref + field.
   *   buildPointer('REF-20250407-001', 'historyJSON')
   *   → '__OVF:REF-20250407-001:historyJSON'
   */
  buildPointer(referenceNumber, fieldName) {
    return OVERFLOW.POINTER_PREFIX + referenceNumber + ':' + fieldName;
  }

  /**
   * Returns true if a cell value is an overflow pointer.
   */
  isPointer(value) {
    return typeof value === 'string' && value.startsWith(OVERFLOW.POINTER_PREFIX);
  }

  /**
   * Parse pointer string back into { referenceNumber, fieldName }.
   * Returns null if the string is not a valid pointer.
   */
  parsePointer(pointer) {
    if (!this.isPointer(pointer)) return null;
    // Strip prefix, split on first colon only
    const body  = pointer.slice(OVERFLOW.POINTER_PREFIX.length);
    const colon = body.indexOf(':');
    if (colon === -1) return null;
    return {
      referenceNumber: body.slice(0, colon),
      fieldName:       body.slice(colon + 1)
    };
  }

  // ── Write path ──────────────────────────────────────────────
  /**
   * write(referenceNumber, fieldName, jsonString)
   *
   * Upserts a row in OverflowStore for this ref+field combination.
   * Returns the pointer string to store in the main Expenses cell.
   *
   * Called by SharedHelper.stringifyJsonFields() when serialized
   * length > OVERFLOW.THRESHOLD.
   */
  write(referenceNumber, fieldName, jsonString) {
    const sheet = this._getSheet();
    const now   = new Date().toISOString();

    // Try to find existing row for this ref+field (O(n) but OverflowStore is small)
    const data   = sheet.getDataRange().getValues();
    const rows   = data.slice(1);  // skip header
    const refCol = OVERFLOW.COL.REF_NUMBER - 1;  // 0-based for array
    const fldCol = OVERFLOW.COL.FIELD_NAME  - 1;

    for (let i = 0; i < rows.length; i++) {
      if (rows[i][refCol] === referenceNumber && rows[i][fldCol] === fieldName) {
        // UPDATE existing row
        const rowNumber = i + 2;  // +1 for header, +1 for 1-based
        sheet.getRange(rowNumber, OVERFLOW.COL.DATA,       1, 1).setValue(jsonString);
        sheet.getRange(rowNumber, OVERFLOW.COL.UPDATED_AT, 1, 1).setValue(now);
        Logger.log('OverflowStore: updated ' + referenceNumber + ':' + fieldName +
                   ' (' + jsonString.length + ' chars)');
        return this.buildPointer(referenceNumber, fieldName);
      }
    }

    // INSERT new row
    sheet.appendRow([referenceNumber, fieldName, jsonString, now]);
    Logger.log('OverflowStore: inserted ' + referenceNumber + ':' + fieldName +
               ' (' + jsonString.length + ' chars)');
    return this.buildPointer(referenceNumber, fieldName);
  }

  // ── Read path ───────────────────────────────────────────────
  /**
   * read(pointer)
   *
   * Given a pointer string, fetches the full JSON string from OverflowStore.
   * Returns null if not found (caller should treat missing data as empty).
   *
   * Called by SharedHelper.parseJsonFields() when cell value isPointer().
   */
  read(pointer) {
    const parsed = this.parsePointer(pointer);
    if (!parsed) return null;

    const sheet = this._getSheet();
    const data  = sheet.getDataRange().getValues();
    const rows  = data.slice(1);
    const refCol = OVERFLOW.COL.REF_NUMBER - 1;
    const fldCol = OVERFLOW.COL.FIELD_NAME  - 1;
    const datCol = OVERFLOW.COL.DATA        - 1;

    for (let i = 0; i < rows.length; i++) {
      if (rows[i][refCol] === parsed.referenceNumber &&
          rows[i][fldCol] === parsed.fieldName) {
        return rows[i][datCol] ? rows[i][datCol].toString() : null;
      }
    }

    Logger.log('WARNING: OverflowStore: pointer not found — ' + pointer);
    return null;
  }

  // ── Cleanup ─────────────────────────────────────────────────
  /**
   * deleteRecord(referenceNumber)
   *
   * Removes ALL overflow rows for a given referenceNumber.
   * Call this if you ever hard-delete an Expenses row (rare).
   * Deletes rows bottom-up to keep row indices valid.
   */
  deleteRecord(referenceNumber) {
    const sheet = this._getSheet();
    const data  = sheet.getDataRange().getValues();
    const refCol = OVERFLOW.COL.REF_NUMBER - 1;

    // Collect matching row numbers (1-based), process bottom-up
    const toDelete = [];
    for (let i = 1; i < data.length; i++) {  // skip header row 0
      if (data[i][refCol] === referenceNumber) {
        toDelete.push(i + 1);  // 1-based sheet row
      }
    }
    toDelete.reverse().forEach(r => sheet.deleteRow(r));

    if (toDelete.length > 0) {
      Logger.log('OverflowStore: deleted ' + toDelete.length + ' rows for ' + referenceNumber);
    }
  }

  // ── Diagnostics ─────────────────────────────────────────────
  /**
   * audit()
   *
   * Returns a summary of all overflow rows.
   * Run from GAS editor for monitoring:
   *   Logger.log(JSON.stringify(_ovf.audit(), null, 2));
   */
  audit() {
    const sheet = this._getSheet();
    const data  = sheet.getDataRange().getValues();
    const rows  = data.slice(1);
    return rows.map(r => ({
      referenceNumber: r[OVERFLOW.COL.REF_NUMBER - 1],
      fieldName:       r[OVERFLOW.COL.FIELD_NAME  - 1],
      charCount:       (r[OVERFLOW.COL.DATA - 1] || '').length,
      updatedAt:       r[OVERFLOW.COL.UPDATED_AT  - 1]
    }));
  }
}

// ── SINGLETON (shared across all service files) ───────────────
// Instantiated lazily to avoid errors if SpreadsheetApp isn't
// available at parse time in some GAS execution contexts.
var _ovfManager = null;

function _getOverflowManager() {
  // GAS is single-threaded per execution context — no race condition possible.
  if (!_ovfManager) {
    _ovfManager = new OverflowManager(SpreadsheetApp.openById(CONFIG.SettingsID));
  }
  return _ovfManager;
}