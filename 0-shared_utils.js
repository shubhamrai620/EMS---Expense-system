/**
 * ============================================================
 *  0-shared_utils.js  —  Shared Utilities & Config
 *  Expense Management System (GAS)
 *
 *  Loaded first (alphabetically). All other files depend on this.
 *  Contains: CONFIG, shared class, and global _wrap helper.
 * ============================================================
 */

// ── MASTER CONFIG ─────────────────────────────────────────────
var CONFIG = {
  NAME:       'Expense Management System',
  INDEX:      'index.html',
  SettingsID: '1LqhOUh1Cy9RpErmHUYS8fKV35-ZxJofbcVIlHZK3UnE',

  // Google Drive folder IDs for file storage
  DRIVE_FOLDERS: {
    EXPENSES: '1tDtGm_c3SlYZb1Z2AG3-sTNfyuTWsVYF',   // Replace with your actual folder ID
    BILLS:    '1B-ProdiX35OfQlFxcZZPrB8DdC66VWqJ'       // Replace with your actual folder ID
  },

  CACHE_TTL: 21600,   // 6 hours in seconds

  SHEET_NAME: {
    USERS:    'Users',
    EXPENSES: 'Expenses'
  },

  // ── Sheet column definitions (single source of truth) ──────
  SHEET_HEADERS: {
    USERS: [
      'UserID', 'name', 'email', 'password', 'role', 'level',
      'jobTitle', 'reportsTo', 'managerName', 'district', 'project', 'projects',
      'department', 'isActive', 'createdDate'
    ],

    // UNIFIED EXPENSES — handles Expenses workflow, plus all payment/bill tracking
    EXPENSES: [
      'referenceNumber',       // REF-YYYYMMDD-NNN (expenses)
      'transactionType',       // EXPENSE
      'ticketNumber',          // TKT-YYYY-NNNN (set on approval)
      'userId',
      'name',
      'email',
      'department',
      'designation',
      'center',
      'expenseType',           // ADVANCE | REIMBURSEMENT (null for requisitions)
      'expenseDate',           // yyyy-MM-dd (for expenses; null for requisitions)
      'status',                // DRAFT | SUBMITTED | APPROVED | REJECTED | CANCELLED
      'ticketStatus',          // CREATED | PAYMENT_PARTIAL | PAYMENT_FULL | BILLED |
                               // BILL_CORRECTION_NEEDED | VERIFIED | BILL_REJECTED |
                               // CLOSED | FORCE_CLOSED | ADMIN_WAIVED | REIMBURSED_CLOSED | SETTLED
      'currentStage',          // MANAGER | OP_HEAD | ACCOUNTS | COMPLETED
      'approvedBy',            // userId of approver (manager or operation head)
      'approvedByName',        // name of approver
      'totalRequestedAmount',
      'totalApprovedAmount',
      'totalBillAmount',
      'totalPaidAmount',
      'expenseItemsJSON',      // ExpenseItem[]
      'bankJSON',              // BankDetails
      'billJSON',              // BillEnvelope  { billSubmissions: BillSubmission[] }
      'paymentJSON',           // PaymentEnvelope { payments: PaymentEntry[] }
      'billVerificationJSON',  // BillVerification[]
      'historyJSON',           // HistoryEntry[]
      'remarksFromOpHead',     // NEW: free-text remarks from OP_HEAD (e.g., rejection reason)
      'remarksFromAccounts',   // free-text notes from accounts team
      'lastUpdatedBy',
      'createdAt',
      'updatedAt',
      'firstPaymentAt',        // timestamp when first payment was made
      'billSubmittedAt',       // timestamp when bills first submitted
      'billVerifiedAt',        // timestamp when bills verified
      'closedAt',              // timestamp when ticket closed
      'settlementType',        // EXCESS | DEFICIT | SETTLED (for ADVANCE tickets after bill verification)
      'settlementAmount',      // Amount to reimburse or recover
      'settledAt',             // Timestamp when settlement was completed
      'mailJSON',              // MailEntry[]  (reserved for future mail/notification system)
      'expenseAttachmentId',   // Google Drive file ID for expense attachment
      'billAttachmentIds'      // Google Drive file IDs for bill attachments (JSON array of file IDs)
    ]
  },

  // ── Status / Stage enumerations ────────────────────────────
  STATUS: {
    DRAFT:              'DRAFT',       // for requisitions before submission
    SUBMITTED:          'SUBMITTED',
    APPROVED:           'APPROVED',
    REJECTED:           'REJECTED',
    CANCELLED:          'CANCELLED',
    HOLD:               'HOLD',        // NEW: Expense on hold by OP_HEAD
    RESUBMIT_REQUIRED:  'RESUBMIT_REQUIRED'  // NEW: Employee must resubmit with corrections
  },

  TICKET_STATUS: {
    CREATED:               'CREATED',
    PAYMENT_PARTIAL:       'PAYMENT_PARTIAL',
    PAYMENT_FULL:          'PAYMENT_FULL',
    PAYMENT_HOLD:          'PAYMENT_HOLD',
    BILLED:                'BILLED',
    BILL_CORRECTION_NEEDED:'BILL_CORRECTION_NEEDED',
    VERIFIED:              'VERIFIED',
    BILL_REJECTED:         'BILL_REJECTED',
    CLOSED:                'CLOSED',
    FORCE_CLOSED:          'FORCE_CLOSED',
    ADMIN_WAIVED:          'ADMIN_WAIVED',
    REIMBURSED_CLOSED:     'REIMBURSED_CLOSED',
    SETTLED:               'SETTLED'    // ADVANCE ticket with bills verified and settled
  },

  STAGE: {
    EMPLOYEE:  'EMPLOYEE',    // NEW: Employee stage for resubmissions
    MANAGER:   'MANAGER',
    OP_HEAD:   'OP_HEAD',      // Operation Head approval stage
    ACCOUNTS:  'ACCOUNTS',
    COMPLETED: 'COMPLETED'
  },

  TRANSACTION_TYPE: {
    EXPENSE: 'EXPENSE'
  },

  EXPENSE_TYPE: {
    ADVANCE:       'ADVANCE',
    REIMBURSEMENT: 'REIMBURSEMENT'
  },

  ROLES: {
    EMPLOYEE:        'employee',
    MANAGER:         'manager',
    OPERATION_HEAD:  'operation_head',
    ACCOUNTS:        'accounts',
    ADMIN:           'admin'
  },

  // Closed terminal states — immutable once reached
  TERMINAL_TICKET_STATUSES: ['CLOSED', 'FORCE_CLOSED', 'ADMIN_WAIVED', 'BILL_REJECTED', 'REIMBURSED_CLOSED', 'SETTLED']
};

// ── VALIDATOR ─────────────────────────────────────────────────
var Validator = {
  isValidEmail: function(e) {
    return e && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e.toString().trim());
  },
  isValidIFSC: function(s) {
    return s && /^[A-Z]{4}0[A-Z0-9]{6}$/.test(s.toString().trim());
  },
  sanitizeString: function(s, max) {
    if (!max) max = 500;
    return s ? s.toString().trim().substring(0, max) : '';
  },
  parseBoolean: function(v) {
    if (typeof v === 'boolean') return v;
    if (typeof v === 'string') {
      var lower = v.toLowerCase();
      return lower === 'true' || v === '1' || lower === 'yes';
    }
    return false;
  },
  requireNonEmpty: function(value, fieldName) {
    if (value === null || value === undefined || value.toString().trim() === '') {
      throw new Error(fieldName + ' is required');
    }
  },
  requirePositiveNumber: function(value, fieldName) {
    var num = parseFloat(value);
    if (isNaN(num) || num <= 0) {
      throw new Error(fieldName + ' must be a positive number');
    }
    return num;
  }
};

// ── SHARED HELPER ─────────────────────────────────────────────
/**
 * SharedHelper — low-level utilities used by all service classes.
 * Instantiated with (db, cache) so it can be unit-tested.
 */
class SharedHelper {
  constructor(db, cache) {
    this.db    = db;
    this.cache = cache;
  }

  // ── Token management ────────────────────────────────────────
  validateToken(token) {
    if (!token) return null;
    const raw = this.cache.get(token);
    if (!raw) return null;
    let parsed;
    try { parsed = JSON.parse(raw); } catch (e) { return null; }
    // Reject stale schema versions
    if (!parsed._v || parsed._v < TOKEN_SCHEMA_VERSION) {
      this.cache.remove(token);
      return null;
    }
    // Sliding expiry: refresh TTL on every valid use
    this.cache.put(token, raw, CONFIG.CACHE_TTL);
    return parsed;
  }

  // ── Sheet helpers ────────────────────────────────────────────
  /**
   * Returns the 0-based column index or throws if missing.
   */
  getHeaderIndex(headers, columnName) {
    const idx = headers.indexOf(columnName);
    if (idx === -1) throw new Error('Column "' + columnName + '" not found in sheet headers');
    return idx;
  }

  /**
   * Zip keys + values into an object. Empty string cells become null.
   */
  createItemObject(keys, values) {
    const obj = {};
    keys.forEach((key, i) => {
      obj[key] = (values[i] !== '' && values[i] !== undefined) ? values[i] : null;
    });
    return obj;
  }

  /**
   * Parse a list of JSON column names in-place on a row object.
   * Falsy / empty cells default to [] or {} based on `arrayFields`.
   */
  parseJsonFields(obj, arrayFields, objectFields) {
    const ovf = _getOverflowManager();   // singleton from 0b-overflow-manager.js
  
    const _resolve = (rawValue) => {
      // If the cell holds a pointer, fetch the real JSON from OverflowStore
      if (ovf.isPointer(rawValue)) {
        const fetched = ovf.read(rawValue);
        return fetched !== null ? fetched : '';   // fall back to empty if missing
      }
      return rawValue;
    };
  
    (arrayFields || []).forEach(f => {
      const raw = _resolve(obj[f]);
      if (raw && raw !== '') {
        try { obj[f] = JSON.parse(raw); } catch (e) { obj[f] = []; }
      } else {
        obj[f] = [];
      }
    });
  
    (objectFields || []).forEach(f => {
      const raw = _resolve(obj[f]);
      if (raw && raw !== '') {
        try { obj[f] = JSON.parse(raw); } catch (e) { obj[f] = {}; }
      } else {
        obj[f] = {};
      }
    });
  
    return obj;
  }
 


  /**
   * Stringify only the JSON columns in a row-data array.
   */
  stringifyJsonFields(obj, jsonFields) {
    const out = Object.assign({}, obj);
    const ovf = _getOverflowManager();   // singleton from 0b-overflow-manager.js
    const ref = obj.referenceNumber || obj.ticketNumber || 'UNKNOWN_REF';
  
    (jsonFields || []).forEach(f => {
      const value = out[f];
  
      if (value !== null && value !== undefined && value !== '') {
        // Serialize to string if not already
        const serialized = (typeof value === 'string') ? value : JSON.stringify(value);
  
        if (serialized.length > OVERFLOW.THRESHOLD) {
          // ── SPILL PATH ──────────────────────────────────────────
          // Write full data to OverflowStore, store pointer in cell
          const pointer = ovf.write(ref, f, serialized);
          out[f] = pointer;
          Logger.log(
            'OVERFLOW: field "' + f + '" on ' + ref +
            ' spilled (' + serialized.length + ' chars → OverflowStore)'
          );
        } else {
          // ── NORMAL PATH ─────────────────────────────────────────
          out[f] = serialized;
        }
      } else {
        out[f] = '';
      }
    });
  
    return out;
  }

  // ── User lookup (no password returned) ─────────────────────
  getUserByEmail(email) {
    email = email.toString().trim().toLowerCase();
    const ws = this.db.getSheetByName(CONFIG.SHEET_NAME.USERS);
    const data = ws.getDataRange().getValues();
    const keys = data[0];
    const rows = data.slice(1);
    const idx  = this.getHeaderIndex(keys, 'email');
    const row  = rows.find(r => (r[idx] || '').toString().trim().toLowerCase() === email);
    if (!row) return null;
    const user = this.createItemObject(keys, row);
    delete user.password;
    return user;
  }

  // Get all users
  getAllUsers() {
    const ws = this.db.getSheetByName(CONFIG.SHEET_NAME.USERS);
    const data = ws.getDataRange().getValues();
    const keys = data[0];
    const rows = data.slice(1);
    return rows.map(row => {
      const user = this.createItemObject(keys, row);
      delete user.password;
      return user;
    });
  }

  // Sequential reference-number generation ──────────────────
  /**
   * Generates the next REF-YYYYMMDD-NNN for Expenses (collision-safe).
   * Reads the last row matching today's date prefix and increments.
   */
  generateExpenseRefNumber(sheet) {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const tz      = Session.getScriptTimeZone();
      const dateStr = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
      const prefix  = 'REF-' + dateStr + '-';

      const [headers, ...rows] = sheet.getDataRange().getValues();
      const refIdx = headers.indexOf('referenceNumber');
      let max = 0;
      rows.forEach(r => {
        const ref = (r[refIdx] || '').toString();
        if (ref.startsWith(prefix)) {
          const num = parseInt(ref.replace(prefix, ''), 10);
          if (!isNaN(num) && num > max) max = num;
        }
      });
      return prefix + String(max + 1).padStart(3, '0');
    } finally {
      lock.releaseLock();
    }
  }

  /**
   * Generates the next TKT-YYYY-NNNN by scanning existing ticketNumber values.
   */
  generateTicketNumber(sheet) {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const year   = new Date().getFullYear().toString();
      const prefix = 'TKT-' + year + '-';
      const [headers, ...rows] = sheet.getDataRange().getValues();
      const tktIdx = headers.indexOf('ticketNumber');
      let max = 0;
      rows.forEach(r => {
        const tkt = (r[tktIdx] || '').toString();
        if (tkt.startsWith(prefix)) {
          const num = parseInt(tkt.replace(prefix, ''), 10);
          if (!isNaN(num) && num > max) max = num;
        }
      });
      return prefix + String(max + 1).padStart(4, '0');
    } finally {
      lock.releaseLock();
    }
  }

  // Drive file upload helpers
  /**
   * Uploads a base64 file to Google Drive and returns the file ID.
   * 
   * FIXED VERSION:
   * - Properly creates blob with MIME type
   * - Sets file sharing permissions
   * - Enhanced error logging
   * 
   * @param {string} base64Data - Base64 encoded file data (WITHOUT data:mime prefix)
   * @param {string} fileName - Name for the file (include extension)
   * @param {string} folderId - Drive folder ID to upload to
   * @param {string} mimeType - MIME type (default: 'application/pdf')
   * @returns {string} Drive file ID
   */
  uploadFileToDrive(base64Data, fileName, folderId, mimeType = 'application/pdf') {
    try {
      // Step 1: Decode base64 to byte array
      const decodedBytes = Utilities.base64Decode(base64Data);
      
      // Step 2: Create blob with proper MIME type
      // The key fix: Utilities.newBlob(data, contentType, name)
      const blob = Utilities.newBlob(decodedBytes, mimeType, fileName);
      
      // Step 3: Get folder and create file
      const folder = DriveApp.getFolderById(folderId);
      const file = folder.createFile(blob);
      
      // Step 4: Set sharing (makes file accessible via link)
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      // Log success
      Logger.log('File uploaded: ' + fileName + '  ' + file.getId());
      
      return file.getId();
    } catch (e) {
      // Enhanced error logging
      Logger.log('Upload failed for: ' + fileName);
      Logger.log('   Error: ' + e.message);
      Logger.log('   Folder ID: ' + folderId);
      Logger.log('   MIME type: ' + mimeType);
      Logger.log('   Base64 length: ' + (base64Data ? base64Data.length : 0));
      
      throw new Error('Failed to upload file to Drive: ' + e.message);
    }
  }

  /**
   * Uploads multiple files to Google Drive and returns array of file IDs.
   * 
   * FIXED VERSION:
   * - Extracts MIME type from each file object
   * - Per-file error handling
   * - Progress logging
   * 
   * @param {Array} files - Array of { name, data, type? } objects (data is base64)
   * @param {string} folderId - Drive folder ID
   * @returns {Array} Array of { name, fileId } objects
   */
  uploadFilesToDrive(files, folderId) {
    if (!files || files.length === 0) return [];
    
    Logger.log('Starting multi-file upload: ' + files.length + ' file(s)');
    
    return files.map((f, index) => {
      try {
        // Extract MIME type from file object or default to PDF
        const mimeType = f.type || 'application/pdf';
        
        // Upload file with proper MIME type
        const fileId = this.uploadFileToDrive(f.data, f.name, folderId, mimeType);
        
        // Log progress
        Logger.log(`Multi-upload [${index + 1}/${files.length}]: ${f.name}  ${fileId}`);
        
        return { name: f.name, fileId: fileId };
      } catch (e) {
        // Log which specific file failed
        Logger.log(`Multi-upload failed [${index + 1}/${files.length}]: ${f.name}`);
        Logger.log('   Error: ' + e.message);
        
        throw new Error(`Failed to upload file "${f.name}": ${e.message}`);
      }
    });
  }

  // ── History helpers ─────────────────────────────────────────
  /**
   * Appends a HistoryEntry to an existing JSON array (string or parsed).
   * Returns a stringified JSON string ready to write back to the sheet.
   *
   * @param {string|Array} existingJSON
   * @param {string} action
   * @param {Object} actor  — { userId, name, role }
   * @param {string} [note]
   * @returns {string}
   */
  addHistoryEntry(existingJSON, action, actor, note) {
    let history = [];
    if (existingJSON && existingJSON !== '') {
      try {
        history = typeof existingJSON === 'string' ? JSON.parse(existingJSON) : existingJSON;
      } catch (e) { history = []; }
    }
    if (!Array.isArray(history)) history = [];
    history.push({
      action,
      by:        actor.userId,
      name:      actor.name,
      role:      actor.role,
      timestamp: new Date().toISOString(),
      note:      note || ''
    });
    return JSON.stringify(history);
  }

  // ── Sheet write helper ──────────────────────────────────────
  /**
   * Writes a full row back to the sheet at the given 1-based rowIndex.
   * `headers` is the ordered list of column names.
   * `obj` is the merged data object; JSON fields are auto-stringified.
   */
  writeRow(sheet, headers, rowIndex, obj, jsonFields) {
    const prepared = this.stringifyJsonFields(obj, jsonFields);
    const rowData  = headers.map(h => {
      const v = prepared[h];
      return (v !== undefined && v !== null) ? v : '';
    });
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowData]);
  }

  // ── Role checks (convenience) ───────────────────────────────
  canApproveExpenses(role) {
    return role === CONFIG.ROLES.OPERATION_HEAD || role === CONFIG.ROLES.ADMIN;
  }
  canProcessPayments(role) {
    return role === CONFIG.ROLES.ACCOUNTS || role === CONFIG.ROLES.ADMIN;
  }
  canVerifyBills(role) {
    return role === CONFIG.ROLES.ACCOUNTS || role === CONFIG.ROLES.ADMIN;
  }
  isAdmin(role) {
    return role === CONFIG.ROLES.ADMIN;
  }
}

// ── GLOBAL WRAP ───────────────────────────────────────────────
/**
 * Wraps any service call in try/catch and returns a JSON string
 * of the form { ok: true, data: ... } or { ok: false, error: { message } }.
 *
 * Usage:
 *   function myRpc(p) { return _wrap(q => _service.myMethod(q))(p); }
 */
function _wrap(fn) {
  return function(params) {
    try {
      var p = (typeof params === 'string') ? JSON.parse(params) : params;
      var result = fn(p);
      if (result && typeof result === 'object' && result.hasOwnProperty('ok')) {
        return JSON.stringify(result);
      }
      return JSON.stringify({ ok: true, data: result });
    } catch (e) {
      return JSON.stringify({ ok: false, error: { message: e.message } });
    }
  };
}