/**
 * ============================================================
 *  2-unified-service.js  —  Unified Transaction Service
 *  Expense Management System (GAS)
 *
 *  FIXED VERSION — All parameter mismatches and field aliases applied
 *
 *  Handles ALL workflows in a single Expenses sheet:
 *    • Expenses (ADVANCE/REIMBURSEMENT)
 *    • Payments
 *    • Bill submission & verification
 *    • Admin overrides
 *    • Accounts dashboard
 *
 *  Depends on: 0-shared_utils.js, 1-backend.js
 * ============================================================
 */

// ── JSON FIELD SCHEMAS ────────────────────────────────────────
//
//  expenseItemsJSON → ExpenseItem[]
//  { head, subHead?, description, amount }
//
//  bankJSON → BankDetails
//  { accountName, accountNumber, ifsc, bankName }
//
//  billJSON → BillEnvelope
//  { billSubmissions: [{ version, totalBillAmount, courierNumber?, courierType?,
//                        attachments: { driveFileId?: string }, items: ExpenseItem[],
//                        submittedAt }] }
//
//  paymentJSON → PaymentEnvelope
//  { payments: [{ tranche, amount, method, reference, date, notes, recordedBy }] }
//
//  billVerificationJSON → BillVerification[]
//  { action, by, name, timestamp, notes, corrections? }
//
//  historyJSON → HistoryEntry[]
//  { action, by, name, role, timestamp, note }
//
//  mailJSON → MailEntry[]  (RESERVED)
//  { type, to, subject, sentAt, status, driveThreadId? }

// ── JSON FIELD LISTS ──────────────────────────────────────────
const JSON_FIELDS  = [
  'expenseItemsJSON', 'bankJSON', 'billJSON',
  'paymentJSON', 'billVerificationJSON', 'historyJSON', 'mailJSON', 'billAttachmentIds'
];
const ARRAY_FIELDS  = ['expenseItemsJSON', 'billVerificationJSON', 'historyJSON', 'mailJSON', 'billAttachmentIds'];;
const OBJECT_FIELDS = ['bankJSON', 'billJSON', 'paymentJSON'];

// ── EXPENSE SHEET HANDLER ─────────────────────────────────────
class ExpenseSheetHandler {
  constructor(db, cache) {
    this.db      = db;
    this.cache   = cache;
    this.sh      = new SharedHelper(db, cache);
    this.sheet   = db.getSheetByName(CONFIG.SHEET_NAME.EXPENSES);
    this.headers = CONFIG.SHEET_HEADERS.EXPENSES;
  }

  // ── Read ────────────────────────────────────────────────────
  getByRef(referenceNumber) {
    const [headers, ...rows] = this.sheet.getDataRange().getValues();
    const refIdx = this.sh.getHeaderIndex(headers, 'referenceNumber');
    const rowIdx = rows.findIndex(r => r[refIdx] === referenceNumber);
    if (rowIdx === -1) return null;

    const obj = this.sh.createItemObject(headers, rows[rowIdx]);
    this.sh.parseJsonFields(obj, ARRAY_FIELDS, OBJECT_FIELDS);

    // Ensure nested structures
    if (!obj.billJSON || typeof obj.billJSON !== 'object')      obj.billJSON    = { billSubmissions: [] };
    if (!obj.billJSON.billSubmissions)                          obj.billJSON.billSubmissions = [];
    if (!obj.paymentJSON || typeof obj.paymentJSON !== 'object') obj.paymentJSON = { payments: [] };
    if (!obj.paymentJSON.payments)                              obj.paymentJSON.payments = [];
    if (!Array.isArray(obj.expenseItemsJSON))                   obj.expenseItemsJSON = [];

    obj._rowIndex = rowIdx + 2;
    return obj;
  }

  getByTicket(ticketNumber) {
    const [headers, ...rows] = this.sheet.getDataRange().getValues();
    const tktIdx = this.sh.getHeaderIndex(headers, 'ticketNumber');
    const rowIdx = rows.findIndex(r => r[tktIdx] === ticketNumber);
    if (rowIdx === -1) return null;

    const obj = this.sh.createItemObject(headers, rows[rowIdx]);
    this.sh.parseJsonFields(obj, ARRAY_FIELDS, OBJECT_FIELDS);
    this._ensureStructures(obj);
    obj._rowIndex = rowIdx + 2;
    return obj;
  }

  getAll() {
    const [headers, ...rows] = this.sheet.getDataRange().getValues();
    return rows.map(row => {
      const obj = this.sh.createItemObject(headers, row);
      this.sh.parseJsonFields(obj, ARRAY_FIELDS, OBJECT_FIELDS);
      this._ensureStructures(obj);
      return obj;
    });
  }

  _ensureStructures(obj) {
    if (!obj.billJSON || typeof obj.billJSON !== 'object')      obj.billJSON    = { billSubmissions: [] };
    if (!obj.paymentJSON || typeof obj.paymentJSON !== 'object') obj.paymentJSON = { payments: [] };
    if (!Array.isArray(obj.expenseItemsJSON))                   obj.expenseItemsJSON = [];
  }

  // ── Write ───────────────────────────────────────────────────
  append(obj) {
    this.sh.writeRow(
      this.sheet, this.headers,
      this.sheet.getLastRow() + 1,
      obj, JSON_FIELDS
    );
  }

  update(referenceNumber, updates, actorUserId) {
    const record = this.getByRef(referenceNumber);
    if (!record) throw new Error('Record not found: ' + referenceNumber);

    Object.assign(record, updates);
    record.lastUpdatedBy = actorUserId;
    record.updatedAt     = new Date().toISOString();

    // Auto-recalculate totals
    if (updates.expenseItemsJSON) {
      record.totalRequestedAmount = _sumItems(record.expenseItemsJSON);
    }
    if (updates.billJSON) {
      record.totalBillAmount = _latestBillTotal(record.billJSON);
    }
    if (updates.paymentJSON) {
      record.totalPaidAmount = _sumPayments(record.paymentJSON);
    }

    this.sh.writeRow(
      this.sheet, this.headers,
      record._rowIndex,
      record, JSON_FIELDS
    );
    return this.getByRef(referenceNumber);
  }

  updateByTicket(ticketNumber, updates) {
    const record = this.getByTicket(ticketNumber);
    if (!record) throw new Error('Ticket not found: ' + ticketNumber);

    Object.assign(record, updates);
    record.updatedAt = new Date().toISOString();

    if (updates.billJSON) {
      record.totalBillAmount = _latestBillTotal(record.billJSON);
    }
    if (updates.paymentJSON) {
      record.totalPaidAmount = _sumPayments(record.paymentJSON);
    }

    this.sh.writeRow(
      this.sheet, this.headers,
      record._rowIndex,
      record, JSON_FIELDS
    );
    return this.getByTicket(ticketNumber);
  }
}

// ── UNIFIED SERVICE ───────────────────────────────────────────
class UnifiedService {
  constructor() {
    this.db      = SpreadsheetApp.openById(CONFIG.SettingsID);
    this.cache   = CacheService.getScriptCache();
    this.sh      = new SharedHelper(this.db, this.cache);
    this.handler = new ExpenseSheetHandler(this.db, this.cache);
  }

  _auth(token) {
    const ud = this.sh.validateToken(token);
    if (!ud) throw new Error('Unauthorized.');
    return ud;
  }

  // ── FIX A: Normalize fields for frontend compatibility ──────
  /**
   * Normalize backend fields to match frontend expectations.
   * Adds field aliases and maps JSON fields to expected names.
   */
  _normalizeForFrontend(r) {
    // ── Status & Type Aliases ──────────────────────────────────
    r.expenseStatus = r.status;              // Frontend expects 'expenseStatus'
    r.sourceType    = r.transactionType;     // Frontend expects 'sourceType'
    
    // ── Amount Aliases ─────────────────────────────────────────
    r.totalAmount = r.totalRequestedAmount || 0;
    r.amount      = r.totalApprovedAmount || 0;      // ✅ For list/table display
    r.paidAmount  = r.totalPaidAmount || 0;          // ✅ For list/table display
    
    // ── Date Fields ────────────────────────────────────────────
    r.expenseDate      = r.expenseDate || r.createdAt;    // ✅ When expense was created
    r.billSubmittedAt  = r.billSubmittedAt || null;       // ✅ When bill was submitted
    
    // ── Employee Info Aliases ──────────────────────────────────
    r.employeeName  = r.name;                // Frontend expects 'employeeName'
    r.employeeEmail = r.email;               // Frontend filters by 'employeeEmail'
    
    // ── Project Code ───────────────────────────────────────────
    r.projectCode = r.project || r.center || '';
    
    // ── JSON FIELD MAPPINGS (CRITICAL!) ────────────────────────
    // Map expenseItemsJSON → items (array of expense line items)
    r.items = Array.isArray(r.expenseItemsJSON) ? r.expenseItemsJSON : [];
    
    // ✅ FIX: Map to billItems for REIMBURSEMENT approval compatibility
    // Frontend expense approval checks for 'billItems' when expenseType is REIMBURSEMENT
    r.billItems = r.items;
    
    // Map bankJSON → bankDetails (bank account info object)
    r.bankDetails = (r.bankJSON && typeof r.bankJSON === 'object') ? r.bankJSON : null;
    
    // Ensure nested bill/payment structures exist
    if (!r.billJSON || typeof r.billJSON !== 'object') {
      r.billJSON = { billSubmissions: [] };
    }
    if (!Array.isArray(r.billJSON.billSubmissions)) {
      r.billJSON.billSubmissions = [];
    }
    
    if (!r.paymentJSON || typeof r.paymentJSON !== 'object') {
      r.paymentJSON = { payments: [] };
    }
    if (!Array.isArray(r.paymentJSON.payments)) {
      r.paymentJSON.payments = [];
    }
    
    // Ensure arrays exist
    if (!Array.isArray(r.billVerificationJSON)) r.billVerificationJSON = [];
    if (!Array.isArray(r.historyJSON))          r.historyJSON = [];
    if (!Array.isArray(r.mailJSON))             r.mailJSON = [];

    // Drive Attachment Fields
    r.expenseAttachmentId = r.expenseAttachmentId || '';

    // Clean Up Internal Fields
    delete r._rowIndex;
    
    return r;
  }
  // ════════════════════════════════════════════════════════════
  //  EXPENSE WORKFLOW (ADVANCE / REIMBURSEMENT)
  // ════════════════════════════════════════════════════════════

  /**
   * createExpense
   * params: { token, expenseType, expenseDate, department?, designation?, center?,
   *           expenseItems[], bankDetails, bills? }
   */
  createExpense({ token, expenseType, expenseDate, department, designation, center,
                  expenseItems, bankDetails, bills,
                  billSubmissionDate, courierNumber, courierType,
                  attachment, billAttachments }) {
    const ud   = this._auth(token);
    const user = this.sh.getUserByEmail(ud.email);
    if (!user) throw new Error('User not found.');

    // Validate
    if (expenseType !== CONFIG.EXPENSE_TYPE.ADVANCE &&
        expenseType !== CONFIG.EXPENSE_TYPE.REIMBURSEMENT) {
      throw new Error('Invalid expense type. Must be ADVANCE or REIMBURSEMENT.');
    }
    if (!expenseDate || !expenseDate.match(/^\d{4}-\d{2}-\d{2}$/)) {
      throw new Error('Invalid expense date. Use yyyy-MM-dd format.');
    }
    if (!Array.isArray(expenseItems) || expenseItems.length === 0) {
      throw new Error('At least one expense item is required.');
    }
    expenseItems.forEach((item, i) => {
      Validator.requirePositiveNumber(item.amount, 'Expense item ' + (i + 1) + ' amount');
    });

    // Bank details are optional - only validate if provided
    if (bankDetails && bankDetails.ifsc && !Validator.isValidIFSC(bankDetails.ifsc)) {
      throw new Error('Invalid IFSC code format.');
    }

    const actorId  = ud.userId || ud.email;
    const refNum   = this.sh.generateExpenseRefNumber(this.handler.sheet);
    const now      = new Date().toISOString();

    // UPLOAD EXPENSE ATTACHMENT
    // Upload expense attachment to Drive if provided
    let expenseAttachmentId = '';
    if (attachment && attachment.data) {
      try {
        // Generate filename with reference number
        const fileName = refNum + '_attachment_' + (attachment.name || 'document.pdf');
        
        // Extract MIME type from attachment object (frontend sends this)
        const mimeType = attachment.type || 'application/pdf';
        
        Logger.log('Uploading expense attachment: ' + fileName);
        Logger.log('   MIME type: ' + mimeType);
        Logger.log('   Size (base64): ' + attachment.data.length + ' chars');
        
        // Upload to Drive with proper MIME type
        expenseAttachmentId = this.sh.uploadFileToDrive(
          attachment.data,           // Base64 data (already stripped of data:mime prefix)
          fileName,                  // Filename with extension
          CONFIG.DRIVE_FOLDERS.EXPENSES,  // Target folder
          mimeType                   // MIME type (e.g., 'application/pdf')
        );
        
        Logger.log('Expense attachment uploaded successfully');
        Logger.log('   File ID: ' + expenseAttachmentId);
      } catch (e) {
        Logger.log('Expense attachment upload FAILED');
        Logger.log('   Error: ' + e.message);
        Logger.log('   Stack: ' + e.stack);
        
        throw new Error('Failed to upload expense attachment: ' + e.message);
      }
    }

    // UPLOAD BILL ATTACHMENTS
    // Upload bill attachments to Drive if provided
    let billAttachmentIds = [];
    if (billAttachments && billAttachments.length > 0) {
      try {
        Logger.log('Uploading bill attachments...');
        Logger.log('   Count: ' + billAttachments.length);
        Logger.log('   Files: ' + billAttachments.map(f => f.name).join(', '));
        
        // uploadFilesToDrive now handles MIME type internally from each file object
        const uploadedFiles = this.sh.uploadFilesToDrive(
          billAttachments,           // Array of { name, data, type }
          CONFIG.DRIVE_FOLDERS.BILLS // Target folder
        );
        
        // Map to name + fileId objects
        billAttachmentIds = uploadedFiles.map(f => ({ 
          name: f.name, 
          fileId: f.fileId 
        }));
        
        Logger.log('Bill attachments uploaded successfully');
        Logger.log('   Count: ' + billAttachmentIds.length);
        billAttachmentIds.forEach((b, i) => {
          Logger.log(`   [${i + 1}] ${b.name}  ${b.fileId}`);
        });
      } catch (e) {
        Logger.log('Bill attachments upload FAILED');
        Logger.log('   Error: ' + e.message);
        Logger.log('   Stack: ' + e.stack);
        
        throw new Error('Failed to upload bill attachments: ' + e.message);
      }
    }

    const billEnvelope = { billSubmissions: [] };
    let billTotal = 0;

    // If creating REIMBURSEMENT with bills attached
    if (expenseType === CONFIG.EXPENSE_TYPE.REIMBURSEMENT && bills && Array.isArray(bills) && bills.length > 0) {
      bills.forEach((b, i) => {
        Validator.requirePositiveNumber(b.amount, 'Bill item ' + (i + 1) + ' amount');
      });
      billTotal = _sumItems(bills);
      billEnvelope.billSubmissions.push({
        version:             1,
        totalBillAmount:     billTotal,
        billSubmissionDate:  billSubmissionDate || now.split('T')[0],  // NEW
        courierNumber:       courierNumber || '',                      // NEW
        courierType:         courierType || '',                        // NEW
        attachments:         {},
        items:               bills.map(item => ({
          ...item,
          expenseDate: item.expenseDate || expenseDate  // NEW: per-item expense date
        })),
        submittedAt:         now
      });
    } else if (expenseType === CONFIG.EXPENSE_TYPE.REIMBURSEMENT) {
      // REIMBURSEMENT without bills initially — still capture submission details for future bills
      billEnvelope.billSubmissions.push({
        version:             1,
        totalBillAmount:     0,
        billSubmissionDate:  billSubmissionDate || now.split('T')[0],  // NEW
        courierNumber:       courierNumber || '',                      // NEW
        courierType:         courierType || '',                        // NEW
        attachments:         {},
        items:               expenseItems.map(item => ({
          ...item,
          expenseDate: item.expenseDate || expenseDate
        })),
        submittedAt:         now
      });
    }

    const total = _sumItems(expenseItems);

    const record = {
      referenceNumber:      refNum,
      transactionType:      CONFIG.TRANSACTION_TYPE.EXPENSE,
      ticketNumber:         '',
      userId:               actorId,
      name:                 user.name,
      email:                user.email,
      department:           department   || user.project || '',
      designation:          designation  || user.jobTitle   || '',
      center:               center       || user.district   || '',
      expenseType,
      expenseDate,
      status:               CONFIG.STATUS.SUBMITTED,
      ticketStatus:         '',
      currentStage:         CONFIG.STAGE.OP_HEAD,
      approvedBy:           '',
      approvedByName:       '',
      totalRequestedAmount: total,
      totalApprovedAmount:  0,
      totalBillAmount:      billTotal,
      totalPaidAmount:      0,
      expenseItemsJSON:     expenseItems,
      bankJSON:             bankDetails,
      billJSON:             billEnvelope,
      paymentJSON:          { payments: [] },
      billVerificationJSON: [],
      historyJSON:          JSON.parse(
                              this.sh.addHistoryEntry('', 'EXPENSE_CREATED',
                                { userId: actorId, name: user.name, role: ud.role },
                                'Expense created (' + expenseType + '). Stage: OP_HEAD')
                            ),
      remarksFromOpHead:    '',
      remarksFromAccounts:  '',
      lastUpdatedBy:        actorId,
      createdAt:            now,
      updatedAt:            now,
      firstPaymentAt:       '',
      billSubmittedAt:      billTotal > 0 ? now : '',
      billVerifiedAt:       '',
      closedAt:             '',
      mailJSON:             [],
      expenseAttachmentId:  expenseAttachmentId,
      billAttachmentIds:    JSON.stringify(billAttachmentIds)
    };

    this.handler.append(record);
    return { referenceNumber: refNum, currentStage: record.currentStage };
  }

  /**
   * getUserExpenses — returns expenses visible to the current user
   * ✅ FIX: Now implements role-based filtering:
   *    - EMPLOYEE: Only their own expenses
   *    - ACCOUNTS: Their own + approved ADVANCE expenses (for payment)
   *    - MANAGER/OP_HEAD/ADMIN: All expenses
   */
  getUserExpenses({ token }) {
    const ud     = this._auth(token);
    const userId = (ud.userId || ud.email).toString();
    const role   = (ud.role || '').toLowerCase();
    
    let expenses = this.handler.getAll()
      .filter(r => r.transactionType === CONFIG.TRANSACTION_TYPE.EXPENSE);
    
    // ── EMPLOYEE: Only their own ─────────────────────────────────
    if (role === CONFIG.ROLES.EMPLOYEE) {
      expenses = expenses.filter(r => r.userId.toString() === userId);
    }
    
    // ACCOUNTS: Their own + approved ADVANCE (for payment)
    else if (role === CONFIG.ROLES.ACCOUNTS) {
      expenses = expenses.filter(r => 
        // Own expenses (if they raise any)
        r.userId.toString() === userId ||
        // OR: Approved ADVANCE expenses needing payment
        (r.expenseType === CONFIG.EXPENSE_TYPE.ADVANCE && 
        r.status === CONFIG.STATUS.APPROVED)
      );
    }

    // MANAGER: See own + direct reports' expenses
    else if (role === 'manager') {
      const allUsers = this.sh.getAllUsers();
      const managerKey = userId.toLowerCase();
      
      // Build Set of userIds: manager's own + all direct reports
      const allowedUserIds = new Set();
      allowedUserIds.add(userId.toString());
      
      allUsers.forEach(u => {
        const reportsTo = (u.reportsTo || '').toString().trim().toLowerCase();
        if (reportsTo === managerKey || reportsTo === ud.email.toLowerCase()) {
          allowedUserIds.add(u.UserID || u.email);
        }
      });
      
      expenses = expenses.filter(r => allowedUserIds.has(r.userId.toString()));
    }
    
    // OP_HEAD: See only expenses from assigned projects ───────────
    else if (role === 'op_head') {
      // Get all assigned projects from both 'project' and 'projects' fields
      const projectField = (ud.project || '').trim();
      const projectsField = (ud.projects || '').trim();
      
      // Build array of assigned projects
      let assignedProjects = [];
      
      // Add from 'project' field (single project)
      if (projectField) {
        assignedProjects.push(projectField);
      }
      
      // Add from 'projects' field (comma-separated)
      if (projectsField) {
        const projectsList = projectsField.split(',')
          .map(p => p.trim())
          .filter(p => p.length > 0);
        assignedProjects = assignedProjects.concat(projectsList);
      }
      
      // Remove duplicates
      assignedProjects = [...new Set(assignedProjects)];
      
      // If Op Head has assigned projects, filter by them
      if (assignedProjects.length > 0) {
        expenses = expenses.filter(r => {
          const expenseProject = (r.project || r.center || '').trim();
          
          // Show if:
          // 1. Expense project matches any assigned project
          // 2. OR it's the Op Head's own expense
          return assignedProjects.includes(expenseProject) || 
                 r.userId.toString() === userId;
        });
        
        Logger.log('OP_HEAD filter applied: ' + ud.name + 
                   ' | Projects: ' + assignedProjects.join(', ') + 
                   ' | Filtered to ' + expenses.length + ' expenses');
      }
      // If no projects assigned, see all (fallback for backward compatibility)
      else {
        Logger.log('OP_HEAD ' + ud.name + ' has no assigned projects - showing all expenses');
      }
    }
    
    // ADMIN: See all ───────────────────────────────────
    // (no filtering needed - already seeing all expenses)
    
    // Debug logging for role-based filtering
    if (role === 'manager' || role === 'op_head') {
      Logger.log('getUserExpenses filter applied:');
      Logger.log('  User: ' + ud.name + ' (' + ud.email + ')');
      Logger.log('  Role: ' + role);
      Logger.log('  Total expenses in system: ' + 
        this.handler.getAll()
          .filter(r => r.transactionType === CONFIG.TRANSACTION_TYPE.EXPENSE)
          .length);
      Logger.log('  Expenses visible to user: ' + expenses.length);
    }
    
    return expenses.map(r => this._normalizeForFrontend(r));
  }


  /**
   * getExpense — fetch single expense by reference
   * ✅ FIX: Now uses _normalizeForFrontend to add expenseStatus alias
   */
  getExpense({ token, referenceNumber }) {
    this._auth(token);
    const exp = this.handler.getByRef(referenceNumber);
    if (!exp) throw new Error('Expense not found.');
    return this._normalizeForFrontend(exp);
  }

  /**
   * forwardExpenseToOpHead — manager forwards to operation head
   */
  forwardExpenseToOpHead({ token, referenceNumber, managerNotes }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    const expense = this.handler.getByRef(referenceNumber);
    if (!expense) throw new Error('Expense not found.');
    if (expense.transactionType !== CONFIG.TRANSACTION_TYPE.EXPENSE) {
      throw new Error('This is not an expense.');
    }
    if (expense.status !== CONFIG.STATUS.SUBMITTED) {
      throw new Error('Expense is not in SUBMITTED state.');
    }
    if (expense.currentStage !== CONFIG.STAGE.MANAGER) {
      throw new Error('Expense is not at MANAGER stage.');
    }

    this.handler.update(referenceNumber, {
      currentStage: CONFIG.STAGE.OP_HEAD,
      historyJSON:  JSON.parse(
                      this.sh.addHistoryEntry(expense.historyJSON, 'FORWARDED_TO_OP_HEAD',
                        { userId, name: user.name, role: ud.role },
                        managerNotes || 'Manager forwarded to Operation Head')
                    )
    }, userId);

    return { currentStage: CONFIG.STAGE.OP_HEAD, message: 'Forwarded to Operation Head.' };
  }

  /**
   * approveExpense — operation head approves expense
   */
  approveExpense({ token, referenceNumber, approvedItems, approverNotes, notes }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    const resolvedNotes = approverNotes || notes || '';

    if (!this.sh.canApproveExpenses(ud.role)) {
      throw new Error('Only Operation Head or Admin can approve expenses.');
    }

    const expense = this.handler.getByRef(referenceNumber);
    if (!expense) throw new Error('Expense not found.');
    if (expense.transactionType !== CONFIG.TRANSACTION_TYPE.EXPENSE) {
      throw new Error('This is not an expense.');
    }
    if (expense.status !== CONFIG.STATUS.SUBMITTED) {
      throw new Error('Expense is not in SUBMITTED state.');
    }

    if (!Array.isArray(approvedItems) || approvedItems.length === 0) {
      throw new Error('At least one approved item is required.');
    }
    approvedItems.forEach((item, i) => {
      Validator.requirePositiveNumber(item.amount, 'Approved item ' + (i + 1) + ' amount');
    });

    const approvedTotal = _sumItems(approvedItems);

    // ✅ NEW LOGIC: Only generate ticket for REIMBURSEMENT
    // For ADVANCE, ticket is generated AFTER first payment
    let ticketNumber = null;
    let ticketStatus = null;

    if (expense.expenseType === CONFIG.EXPENSE_TYPE.REIMBURSEMENT) {
      ticketNumber = this.sh.generateTicketNumber(this.handler.sheet);
      // ✅ FIX: Set to BILLED if bills are already submitted, otherwise CREATED
      const hasBills = expense.billJSON && Array.isArray(expense.billJSON.billSubmissions) && expense.billJSON.billSubmissions.length > 0;
      ticketStatus = hasBills ? CONFIG.TICKET_STATUS.BILLED : CONFIG.TICKET_STATUS.CREATED;
    }

    this.handler.update(referenceNumber, {
      status:              CONFIG.STATUS.APPROVED,
      currentStage:        CONFIG.STAGE.ACCOUNTS,
      ticketNumber:        ticketNumber || '',  // Empty for ADVANCE
      ticketStatus:        ticketStatus || '',  // Empty for ADVANCE
      approvedBy:          userId,
      approvedByName:      user.name,
      totalApprovedAmount: approvedTotal,
      expenseItemsJSON:    approvedItems,
      historyJSON:         JSON.parse(
                            this.sh.addHistoryEntry(expense.historyJSON, 'APPROVED_BY_OP_HEAD',
                              { userId, name: user.name, role: ud.role },
                              'Approved: ₹' + approvedTotal + '. ' + resolvedNotes)
                          )
    }, userId);

    return {
      ticketNumber:   ticketNumber || '',
      ticketStatus:   ticketStatus || '',
      expenseType:    expense.expenseType,  // ✅ NEW: Return expense type
      message:        ticketNumber 
        ? 'Expense approved. Ticket ' + ticketNumber + ' created.'
        : 'Expense approved. Awaiting payment processing.'
    };
  }

  /**
   * processExpensePayment — accounts processes payment for ADVANCE expenses
   * This is called BEFORE ticket generation (only for ADVANCE type)
   */
  processExpensePayment({ token, referenceNumber, amount, method, paymentMethod, reference, paymentDate, notes }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    if (!this.sh.canProcessPayments(ud.role)) {
      throw new Error('Only Accounts or Admin can process payments.');
    }

    const expense = this.handler.getByRef(referenceNumber);
    if (!expense) throw new Error('Expense not found.');
    if (expense.transactionType !== CONFIG.TRANSACTION_TYPE.EXPENSE) {
      throw new Error('This is not an expense.');
    }
    if (expense.status !== CONFIG.STATUS.APPROVED) {
      throw new Error('Expense must be APPROVED before processing payment.');
    }
    if (expense.expenseType !== CONFIG.EXPENSE_TYPE.ADVANCE) {
      throw new Error('Only ADVANCE expenses can have payments processed here.');
    }

    Validator.requirePositiveNumber(amount, 'Payment amount');

    const approved = parseFloat(expense.totalApprovedAmount) || 0;
    const paid     = parseFloat(expense.totalPaidAmount) || 0;
    const newTotal = paid + parseFloat(amount);
    const resolvedMethod = method || paymentMethod || '';

    if (newTotal > approved) {
      throw new Error('Total payment (₹' + newTotal + ') exceeds approved amount (₹' + approved + ').');
    }

    // Initialize or get existing payment envelope
    let envelope = expense.paymentJSON;
    if (!envelope || !envelope.payments) {
      envelope = { payments: [] };
    }
    const tranche = envelope.payments.length + 1;

    envelope.payments.push({
      tranche,
      amount:     parseFloat(amount),
      method:     resolvedMethod,
      reference:  reference  || '',
      date:       paymentDate || new Date().toISOString().split('T')[0],
      notes:      notes      || '',
      recordedBy: user.name
    });

    // Determine new ticket status
    const isFullPayment = (newTotal >= approved);
    let ticketNumber = expense.ticketNumber;
    let ticketStatus = expense.ticketStatus;

    // ✅ Generate ticket on FIRST payment (if not already generated)
    if (!ticketNumber) {
      ticketNumber = this.sh.generateTicketNumber(this.handler.sheet);
      ticketStatus = isFullPayment 
        ? CONFIG.TICKET_STATUS.PAYMENT_FULL 
        : CONFIG.TICKET_STATUS.PAYMENT_PARTIAL;
    } else {
      // Update status for subsequent payments
      ticketStatus = isFullPayment 
        ? CONFIG.TICKET_STATUS.PAYMENT_FULL 
        : CONFIG.TICKET_STATUS.PAYMENT_PARTIAL;
    }

    const now = new Date().toISOString();
    this.handler.update(referenceNumber, {
      paymentJSON:        envelope,
      totalPaidAmount:    newTotal,
      ticketNumber,
      ticketStatus,
      firstPaymentAt:     expense.firstPaymentAt || now,
      historyJSON:        JSON.parse(
                            this.sh.addHistoryEntry(expense.historyJSON, 'PAYMENT_RECORDED',
                              { userId, name: user.name, role: ud.role },
                              'Payment #' + tranche + ': ₹' + amount + '. Total paid: ₹' + newTotal)
                          )
    }, userId);

    return {
      tranche,
      totalPaidAmount: newTotal,
      remaining:       Math.max(0, approved - newTotal),
      ticketNumber,
      ticketStatus,
      message:         'Payment recorded.' + (tranche === 1 ? ' Ticket ' + ticketNumber + ' created.' : '')
    };
  }

  /**
   * rejectExpense
   * ✅ FIX B: Now accepts both 'rejectionReason' and 'reason' (frontend sends 'reason')
   */
  rejectExpense({ token, referenceNumber, rejectionReason, reason }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    if (!this.sh.canApproveExpenses(ud.role)) {
      throw new Error('Only Operation Head or Admin can reject expenses.');
    }

    const resolvedReason = rejectionReason || reason || '';
    Validator.requireNonEmpty(resolvedReason, 'Rejection reason');

    const expense = this.handler.getByRef(referenceNumber);
    if (!expense) throw new Error('Expense not found.');
    if (expense.transactionType !== CONFIG.TRANSACTION_TYPE.EXPENSE) {
      throw new Error('This is not an expense.');
    }
    if (expense.status !== CONFIG.STATUS.SUBMITTED) {
      throw new Error('Expense is not in SUBMITTED state.');
    }

    this.handler.update(referenceNumber, {
      status:            CONFIG.STATUS.RESUBMIT_REQUIRED,  // NEW: Instead of REJECTED
      currentStage:      CONFIG.STAGE.EMPLOYEE,            // NEW: Send back to employee
      remarksFromOpHead: resolvedReason,                   // NEW: Store rejection remarks
      historyJSON:       JSON.parse(
                           this.sh.addHistoryEntry(expense.historyJSON, 'REJECTED_BY_OP_HEAD',
                             { userId, name: user.name, role: ud.role },
                             'Rejected. Reason: ' + resolvedReason)
                         )
    }, userId);

    try {
      _mailExpenseRejected(expense, resolvedReason);
    } catch(e) {
      Logger.log('Mail stub failed: ' + e.message);
    }
    return { status: CONFIG.STATUS.RESUBMIT_REQUIRED, message: 'Expense rejected. Employee can resubmit with corrections.' };
  }

  /**
   * cancelExpense — employee can cancel before approval
   * ✅ FIX C: Now accepts both 'cancellationReason' and 'reason'
   */
  cancelExpense({ token, referenceNumber, cancellationReason, reason }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    const expense = this.handler.getByRef(referenceNumber);
    if (!expense) throw new Error('Expense not found.');
    if (expense.transactionType !== CONFIG.TRANSACTION_TYPE.EXPENSE) {
      throw new Error('This is not an expense.');
    }
    if (expense.userId.toString() !== userId.toString() && !this.sh.isAdmin(ud.role)) {
      throw new Error('Only the expense owner or Admin can cancel.');
    }
    if (expense.status !== CONFIG.STATUS.SUBMITTED) {
      throw new Error('Expense cannot be cancelled in its current state.');
    }

    const resolvedReason = cancellationReason || reason || '';

    this.handler.update(referenceNumber, {
      status:       CONFIG.STATUS.CANCELLED,
      currentStage: CONFIG.STAGE.COMPLETED,
      historyJSON:  JSON.parse(
                      this.sh.addHistoryEntry(expense.historyJSON, 'CANCELLED',
                        { userId, name: user.name, role: ud.role },
                        resolvedReason || 'Cancelled by user')
                    )
    }, userId);

    return { status: CONFIG.STATUS.CANCELLED, message: 'Expense cancelled.' };
  }

  /**
   * holdExpense — OP_HEAD puts expense on hold pending review
   */
  holdExpense({ token, referenceNumber, reason }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    if (!this.sh.canApproveExpenses(ud.role)) {
      throw new Error('Only Operation Head or Admin can hold expenses.');
    }

    const expense = this.handler.getByRef(referenceNumber);
    if (!expense) throw new Error('Expense not found.');
    if (expense.status !== CONFIG.STATUS.SUBMITTED) {
      throw new Error('Only SUBMITTED expenses can be held.');
    }

    this.handler.update(referenceNumber, {
      status:      CONFIG.STATUS.HOLD,
      historyJSON: JSON.parse(
                     this.sh.addHistoryEntry(expense.historyJSON, 'HOLD',
                       { userId, name: user.name, role: ud.role },
                       reason || 'Expense put on hold')
                   )
    }, userId);

    return { success: true, message: 'Expense held successfully' };
  }

  /**
   * releaseExpenseHold — OP_HEAD releases hold, expense back to SUBMITTED
   */
  releaseExpenseHold({ token, referenceNumber }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    if (!this.sh.canApproveExpenses(ud.role)) {
      throw new Error('Only Operation Head or Admin can release holds.');
    }

    const expense = this.handler.getByRef(referenceNumber);
    if (!expense) throw new Error('Expense not found.');
    if (expense.status !== CONFIG.STATUS.HOLD) {
      throw new Error('Expense is not on hold.');
    }

    this.handler.update(referenceNumber, {
      status:      CONFIG.STATUS.SUBMITTED,
      historyJSON: JSON.parse(
                     this.sh.addHistoryEntry(expense.historyJSON, 'HOLD_RELEASED',
                       { userId, name: user.name, role: ud.role },
                       'Hold released. Expense back in review queue.')
                   )
    }, userId);

    return { success: true, message: 'Hold released successfully' };
  }

  /**
   * resubmitExpense — Employee resubmits expense after RESUBMIT_REQUIRED rejection
   */
  resubmitExpense({ token, referenceNumber, expenseItems }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    const expense = this.handler.getByRef(referenceNumber);
    if (!expense) throw new Error('Expense not found.');

    // Verify ownership
    if (expense.userId.toString() !== userId.toString() && !this.sh.isAdmin(ud.role)) {
      throw new Error('Unauthorized: Can only resubmit your own expenses');
    }

    // Verify status
    if (expense.status !== CONFIG.STATUS.RESUBMIT_REQUIRED) {
      throw new Error('Expense is not in RESUBMIT_REQUIRED state. Current status: ' + expense.status);
    }

    // Validate items
    if (!Array.isArray(expenseItems) || expenseItems.length === 0) {
      throw new Error('At least one expense item is required.');
    }
    expenseItems.forEach((item, i) => {
      Validator.requirePositiveNumber(item.amount, 'Expense item ' + (i + 1) + ' amount');
    });

    // Calculate new total
    const totalAmount = _sumItems(expenseItems);

    this.handler.update(referenceNumber, {
      expenseItemsJSON:    expenseItems,
      totalRequestedAmount: totalAmount,
      totalApprovedAmount:  0,
      status:              CONFIG.STATUS.SUBMITTED,
      currentStage:        CONFIG.STAGE.OP_HEAD,
      remarksFromOpHead:   '',  // Clear rejection remarks
      historyJSON:         JSON.parse(
                             this.sh.addHistoryEntry(expense.historyJSON, 'RESUBMITTED',
                               { userId, name: user.name, role: ud.role },
                               'Expense resubmitted after corrections. New amount: ₹' + totalAmount.toLocaleString('en-IN'))
                           )
    }, userId);

    return {
      success: true,
      message: 'Expense resubmitted successfully',
      referenceNumber: referenceNumber,
      totalAmount: totalAmount
    };
  }

  /**
   * updateExpenseItems — edit items before approval
   */
  updateExpenseItems({ token, referenceNumber, expenseItems }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;

    const expense = this.handler.getByRef(referenceNumber);
    if (!expense) throw new Error('Expense not found.');
    if (expense.transactionType !== CONFIG.TRANSACTION_TYPE.EXPENSE) {
      throw new Error('This is not an expense.');
    }
    if (expense.userId.toString() !== userId.toString()) {
      throw new Error('Only the expense owner can update items.');
    }
    if (expense.status !== CONFIG.STATUS.SUBMITTED) {
      throw new Error('Cannot update items after approval/rejection.');
    }

    if (!Array.isArray(expenseItems) || expenseItems.length === 0) {
      throw new Error('At least one expense item is required.');
    }
    expenseItems.forEach((item, i) => {
      Validator.requirePositiveNumber(item.amount, 'Item ' + (i + 1) + ' amount');
    });

    this.handler.update(referenceNumber, {
      expenseItemsJSON: expenseItems,
      historyJSON:      JSON.parse(
                          this.sh.addHistoryEntry(expense.historyJSON, 'ITEMS_UPDATED',
                            { userId, name: expense.name, role: ud.role },
                            'Expense items updated')
                        )
    }, userId);

    return { message: 'Expense items updated.' };
  }

  /**
   * updateBankDetails — edit bank info before approval
   */
  updateBankDetails({ token, referenceNumber, bankDetails }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;

    const expense = this.handler.getByRef(referenceNumber);
    if (!expense) throw new Error('Expense not found.');
    if (expense.transactionType !== CONFIG.TRANSACTION_TYPE.EXPENSE) {
      throw new Error('This is not an expense.');
    }
    if (expense.userId.toString() !== userId.toString()) {
      throw new Error('Only the expense owner can update bank details.');
    }
    if (expense.status !== CONFIG.STATUS.SUBMITTED) {
      throw new Error('Cannot update bank details after approval/rejection.');
    }

    // Bank details are optional - only validate if provided
    if (bankDetails && bankDetails.ifsc && !Validator.isValidIFSC(bankDetails.ifsc)) {
      throw new Error('Invalid IFSC code format.');
    }

    this.handler.update(referenceNumber, {
      bankJSON:     bankDetails,
      historyJSON:  JSON.parse(
                      this.sh.addHistoryEntry(expense.historyJSON, 'BANK_DETAILS_UPDATED',
                        { userId, name: expense.name, role: ud.role },
                        'Bank details updated')
                    )
    }, userId);

    return { message: 'Bank details updated.' };
  }

  /**
   * submitBill — employee submits bills for reimbursement expense
   */
  submitBill({ token, referenceNumber, items, courierNumber, courierType, attachments }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;

    const expense = this.handler.getByRef(referenceNumber);
    if (!expense) throw new Error('Expense not found.');
    if (expense.transactionType !== CONFIG.TRANSACTION_TYPE.EXPENSE) {
      throw new Error('This is not an expense.');
    }
    if (expense.userId.toString() !== userId.toString()) {
      throw new Error('Only the expense owner can submit bills.');
    }
    if (expense.expenseType !== CONFIG.EXPENSE_TYPE.REIMBURSEMENT) {
      throw new Error('Bills can only be submitted for REIMBURSEMENT expenses.');
    }

    if (!Array.isArray(items) || items.length === 0) {
      throw new Error('At least one bill item is required.');
    }
    items.forEach((b, i) => {
      Validator.requirePositiveNumber(b.amount, 'Bill item ' + (i + 1) + ' amount');
    });

    const billTotal = _sumItems(items);
    const envelope  = expense.billJSON;
    const version   = envelope.billSubmissions.length + 1;
    const now       = new Date().toISOString();

    envelope.billSubmissions.push({
      version,
      totalBillAmount: billTotal,
      courierNumber:   courierNumber || '',
      courierType:     courierType   || '',
      attachments:     attachments   || {},
      items,
      submittedAt:     now
    });

    this.handler.update(referenceNumber, {
      billJSON:        envelope,
      billSubmittedAt: expense.billSubmittedAt || now,
      historyJSON:     JSON.parse(
                         this.sh.addHistoryEntry(expense.historyJSON, 'BILLS_SUBMITTED',
                           { userId, name: expense.name, role: ud.role },
                           'Bill version ' + version + ' submitted: ₹' + billTotal)
                       )
    }, userId);

    return { billVersion: version, totalBillAmount: billTotal, message: 'Bills submitted.' };
  }

  /**
   * getExpenseStats — summary stats for the user
   */
  getExpenseStats({ token }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;

    const rows = this.handler.getAll()
      .filter(r => r.userId === userId && r.transactionType === CONFIG.TRANSACTION_TYPE.EXPENSE);

    let submitted = 0, approved = 0, rejected = 0, cancelled = 0,
        totalRequested = 0, totalApproved = 0;

    rows.forEach(r => {
      if (r.status === CONFIG.STATUS.SUBMITTED)  submitted++;
      if (r.status === CONFIG.STATUS.APPROVED)   approved++;
      if (r.status === CONFIG.STATUS.REJECTED)   rejected++;
      if (r.status === CONFIG.STATUS.CANCELLED)  cancelled++;

      totalRequested += parseFloat(r.totalRequestedAmount) || 0;
      if (r.status === CONFIG.STATUS.APPROVED) {
        totalApproved += parseFloat(r.totalApprovedAmount) || 0;
      }
    });

    return {
      submitted, approved, rejected, cancelled,
      totalRequested, totalApproved,
      total: rows.length
    };
  }

  /**
   * getPendingExpenseApprovals — for operation head
   * ✅ FIX: Now uses _normalizeForFrontend to add expenseStatus alias
   */
  getPendingExpenseApprovals({ token }) {
    const ud = this._auth(token);
    if (!this.sh.canApproveExpenses(ud.role)) {
      throw new Error('Only Operation Head or Admin can view pending approvals.');
    }

    return this.handler.getAll()
      .filter(r => r.transactionType === CONFIG.TRANSACTION_TYPE.EXPENSE && r.status === CONFIG.STATUS.SUBMITTED)
      .map(r => this._normalizeForFrontend(r));
  }

  /**
   * getTickets — view tickets (approved expenses/requisitions with ticketNumber)
   */
  getTickets({ token, ticketStatus, filters }) {
    const ud     = this._auth(token);
    const userId = (ud.userId || ud.email).toString();
    const role   = (ud.role || '').toLowerCase();

    const statusFilter = ticketStatus || (filters && filters.ticketStatus) || '';

    let tickets = this.handler.getAll()
      .filter(r => r.ticketNumber && r.ticketNumber !== '');

    if (role === CONFIG.ROLES.EMPLOYEE) {
      tickets = tickets.filter(t => t.userId.toString() === userId);
    } else if (!this.sh.canProcessPayments(ud.role)) {
      throw new Error('Unauthorized to view tickets.');
    }

    if (statusFilter) {
      tickets = tickets.filter(t => t.ticketStatus === statusFilter);
    }

    return tickets.map(t => this._normalizeForFrontend(t));
  }

  /**
   * getTicketDetail
   */
  getTicketDetail({ token, ticketNumber }) {
    const ud = this._auth(token);
    const ticket = this.handler.getByTicket(ticketNumber);
    if (!ticket) throw new Error('Ticket not found.');

    const userId  = ud.userId || ud.email;
    const isOwner = ticket.userId === userId;
    if (!isOwner && !this.sh.canProcessPayments(ud.role)) {
      throw new Error('Only the ticket owner, Accounts, or Admin can view ticket details.');
    }

    delete ticket._rowIndex;
    return ticket;
  }

  /**
   * getPendingPayments — tickets awaiting payment
   */
  getPendingPayments({ token }) {
    const ud = this._auth(token);
    if (!this.sh.canProcessPayments(ud.role)) {
      throw new Error('Only Accounts or Admin can view pending payments.');
    }

    const S = CONFIG.TICKET_STATUS;
    return this.handler.getAll()
      .filter(r => r.ticketNumber && (r.ticketStatus === S.CREATED || r.ticketStatus === S.PAYMENT_PARTIAL || r.ticketStatus === S.PAYMENT_HOLD))
      .map(r => { delete r._rowIndex; return r; });
  }

  /**
   * getPendingBillVerification — tickets awaiting bill verification
   */
  getPendingBillVerification({ token }) {
    const ud = this._auth(token);
    if (!this.sh.canProcessPayments(ud.role)) {
      throw new Error('Only Accounts or Admin can view pending bill verification.');
    }

    const S = CONFIG.TICKET_STATUS;
    return this.handler.getAll()
      .filter(r => r.ticketNumber && (r.ticketStatus === S.BILLED || r.ticketStatus === S.BILL_CORRECTION_NEEDED))
      .map(r => { delete r._rowIndex; return r; });
  }

  /**
   * processPayment — accounts records a payment
   * ✅ FIX E: Now accepts both 'method' and 'paymentMethod', returns 'remaining' field
   * ✅ FIX F: Now accepts billDate parameter
   */
  processPayment({ token, ticketNumber, amount, method, paymentMethod, reference, transactionRef, paymentDate, billDate, notes }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    if (!this.sh.canProcessPayments(ud.role)) {
      throw new Error('Only Accounts or Admin can process payments.');
    }

    const ticket = this.handler.getByTicket(ticketNumber);
    if (!ticket) throw new Error('Ticket not found.');

    if (CONFIG.TERMINAL_TICKET_STATUSES.includes(ticket.ticketStatus)) {
      throw new Error('Ticket is in terminal state: ' + ticket.ticketStatus);
    }

    Validator.requirePositiveNumber(amount, 'Payment amount');

    const approved     = parseFloat(ticket.totalApprovedAmount) || 0;
    const paid         = parseFloat(ticket.totalPaidAmount)     || 0;
    const newTotal     = paid + parseFloat(amount);
    const resolvedMethod = method || paymentMethod || '';
    const resolvedRef = reference || transactionRef || '';

    if (newTotal > approved) {
      throw new Error('Total payment (₹' + newTotal + ') exceeds approved amount (₹' + approved + ').');
    }

    const envelope = ticket.paymentJSON;
    const tranche  = envelope.payments.length + 1;

    envelope.payments.push({
      tranche,
      amount:     parseFloat(amount),
      method:     resolvedMethod,
      reference:  resolvedRef,
      date:       paymentDate || new Date().toISOString().split('T')[0],
      notes:      notes      || '',
      recordedBy: user.name,
      billDate:   billDate   || ''  // ✅ Store bill date if provided
    });

    const newStatus = (newTotal >= approved)
      ? CONFIG.TICKET_STATUS.PAYMENT_FULL
      : CONFIG.TICKET_STATUS.PAYMENT_PARTIAL;

    const now = new Date().toISOString();
    this.handler.updateByTicket(ticketNumber, {
      paymentJSON:     envelope,
      ticketStatus:    newStatus,
      billDate:        billDate || ticket.billDate,  // ✅ Also store at top level
      firstPaymentAt:  ticket.firstPaymentAt || now,
      historyJSON:     JSON.parse(
                         this.sh.addHistoryEntry(ticket.historyJSON, 'PAYMENT_RECORDED',
                           { userId, name: user.name, role: ud.role },
                           'Payment #' + tranche + ': ₹' + amount + '. Total paid: ₹' + newTotal)
                       )
    });

    return {
      tranche,
      totalPaidAmount: newTotal,
      remaining:       Math.max(0, approved - newTotal),  // ✅ FIX: add remaining field
      ticketStatus:    newStatus,
      message:         'Payment recorded.'
    };
  }

  /**
   * flagPaymentDelay — accounts flags a payment issue
   * ✅ FIX F: Now accepts both 'delayNotes' and 'notes' (frontend sends 'notes')
   */
  flagPaymentDelay({ token, ticketNumber, delayNotes, notes }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    if (!this.sh.canProcessPayments(ud.role)) {
      throw new Error('Only Accounts or Admin can flag payment delays.');
    }

    const resolvedNotes = delayNotes || notes || '';
    Validator.requireNonEmpty(resolvedNotes, 'Delay notes');

    const ticket = this.handler.getByTicket(ticketNumber);
    if (!ticket) throw new Error('Ticket not found.');

    this.handler.updateByTicket(ticketNumber, {
      remarksFromAccounts: resolvedNotes,
      historyJSON:         JSON.parse(
                             this.sh.addHistoryEntry(ticket.historyJSON, 'PAYMENT_DELAY_FLAGGED',
                               { userId, name: user.name, role: ud.role },
                               'Payment delayed: ' + resolvedNotes)
                           )
    });

    try {
      _mailPaymentDelayed(ticket, resolvedNotes);
    } catch(e) {
      Logger.log('Mail stub failed: ' + e.message);
    }
    return { message: 'Payment delay flagged.' };
  }

  /**
   * holdPayment - accounts puts a payment on hold with mandatory reason
   */
  holdPayment({ token, ticketNumber, holdReason }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    if (!this.sh.canProcessPayments(ud.role)) {
      throw new Error('Only Accounts or Admin can hold payments.');
    }

    Validator.requireNonEmpty(holdReason, 'Hold reason');

    const ticket = this.handler.getByTicket(ticketNumber);
    if (!ticket) throw new Error('Ticket not found.');

    if (ticket.ticketStatus !== CONFIG.TICKET_STATUS.CREATED &&
        ticket.ticketStatus !== CONFIG.TICKET_STATUS.PAYMENT_PARTIAL) {
      throw new Error('Payment can only be put on hold when ticket status is CREATED or PAYMENT_PARTIAL.');
    }

    this.handler.updateByTicket(ticketNumber, {
      ticketStatus:        CONFIG.TICKET_STATUS.PAYMENT_HOLD,
      remarksFromAccounts: holdReason,
      historyJSON:         JSON.parse(
                            this.sh.addHistoryEntry(ticket.historyJSON, 'PAYMENT_HOLD',
                              { userId, name: user.name, role: ud.role },
                              'Payment held: ' + holdReason)
                          )
    });

    try {
      _mailPaymentDelayed(ticket, holdReason);
    } catch(e) {
      Logger.log('Mail stub failed: ' + e.message);
    }
    return { ticketStatus: CONFIG.TICKET_STATUS.PAYMENT_HOLD, message: 'Payment placed on hold.' };
  }

  /**
   * releasePaymentHold - accounts releases a payment hold
   */
  releasePaymentHold({ token, ticketNumber, releaseNotes }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    if (!this.sh.canProcessPayments(ud.role)) {
      throw new Error('Only Accounts or Admin can release payment holds.');
    }

    const ticket = this.handler.getByTicket(ticketNumber);
    if (!ticket) throw new Error('Ticket not found.');

    if (ticket.ticketStatus !== CONFIG.TICKET_STATUS.PAYMENT_HOLD) {
      throw new Error('Ticket is not on payment hold.');
    }

    const paidAmount = parseFloat(ticket.totalPaidAmount) || 0;
    const restoredStatus = paidAmount > 0 
      ? CONFIG.TICKET_STATUS.PAYMENT_PARTIAL 
      : CONFIG.TICKET_STATUS.CREATED;

    this.handler.updateByTicket(ticketNumber, {
      ticketStatus:        restoredStatus,
      remarksFromAccounts: releaseNotes || '',
      historyJSON:         JSON.parse(
                            this.sh.addHistoryEntry(ticket.historyJSON, 'PAYMENT_HOLD_RELEASED',
                              { userId, name: user.name, role: ud.role },
                              'Payment hold released' + (releaseNotes ? ': ' + releaseNotes : ''))
                          )
    });

    return { ticketStatus: restoredStatus, message: 'Payment hold released.' };
  }

  /**
   * submitBillsOnTicket — employee submits bills via ticket
   */
  submitBillsOnTicket({ token, ticketNumber, items, courierNumber, courierType, attachments }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    const ticket = this.handler.getByTicket(ticketNumber);
    if (!ticket) throw new Error('Ticket not found.');

    if (ticket.userId.toString() !== userId.toString() && !this.sh.isAdmin(ud.role)) {
      throw new Error('Only the expense owner can submit bills.');
    }

    if (CONFIG.TERMINAL_TICKET_STATUSES.includes(ticket.ticketStatus)) {
      throw new Error('Ticket is in terminal state: ' + ticket.ticketStatus);
    }

    // Prevent bill submission for REIMBURSEMENT expense types
    // Reimbursement expenses submit bills during expense creation, not via ticket panel
    if (ticket.expenseType === CONFIG.EXPENSE_TYPE.REIMBURSEMENT) {
      throw new Error('Bills for reimbursement expenses must be submitted during expense creation, not via ticket panel.');
    }

    if (!Array.isArray(items) || items.length === 0) {
      throw new Error('At least one bill item is required.');
    }
    items.forEach((b, i) => {
      Validator.requirePositiveNumber(b.amount, 'Bill item ' + (i + 1) + ' amount');
    });

    const billTotal = _sumItems(items);
    const envelope  = ticket.billJSON;
    const version   = envelope.billSubmissions.length + 1;
    const now       = new Date().toISOString();

    envelope.billSubmissions.push({
      version,
      totalBillAmount: billTotal,
      courierNumber:   courierNumber || '',
      courierType:     courierType   || '',
      attachments:     attachments   || {},
      items,
      submittedAt:     now
    });

    this.handler.updateByTicket(ticketNumber, {
      billJSON:        envelope,
      ticketStatus:    CONFIG.TICKET_STATUS.BILLED,
      billSubmittedAt: ticket.billSubmittedAt || now,
      historyJSON:     JSON.parse(
                         this.sh.addHistoryEntry(ticket.historyJSON, 'BILLS_SUBMITTED',
                           { userId, name: user.name, role: ud.role },
                           'Bills submitted — version ' + version + ': ₹' + billTotal)
                       )
    });

    return { billVersion: version, totalBillAmount: billTotal, message: 'Bills submitted.' };
  }

  /**
   * approveBills — accounts approves bills
   */
  approveBills({ token, ticketNumber, verificationNotes }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    if (!this.sh.canVerifyBills(ud.role)) {
      throw new Error('Only Accounts or Admin can verify bills.');
    }

    const ticket = this.handler.getByTicket(ticketNumber);
    if (!ticket) throw new Error('Ticket not found.');

    if (ticket.ticketStatus !== CONFIG.TICKET_STATUS.BILLED &&
        ticket.ticketStatus !== CONFIG.TICKET_STATUS.BILL_CORRECTION_NEEDED) {
      throw new Error('Ticket is not in BILLED state.');
    }

    const envelope = ticket.billVerificationJSON || [];
    envelope.push({
      action:    'APPROVED',
      by:        userId,
      name:      user.name,
      timestamp: new Date().toISOString(),
      notes:     verificationNotes || 'Bills approved'
    });

    const now = new Date().toISOString();
    
    // Calculate settlement for ADVANCE expenses
    let settlementType = null;
    let settlementAmount = 0;
    
    // Get the latest bill submission to calculate total
    const latestBill = ticket.billJSON?.billSubmissions?.[ticket.billJSON.billSubmissions.length - 1];
    let billTotalAmount = 0;
    if (latestBill && latestBill.items) {
      billTotalAmount = (latestBill.items || []).reduce((sum, item) => sum + (parseFloat(item.amount) || 0), 0);
    }
    
    const advanceAmount = parseFloat(ticket.amount) || 0;
    const diff = billTotalAmount - advanceAmount;
    
    if (ticket.expenseType === CONFIG.EXPENSE_TYPE.ADVANCE) {
      // Calculate settlement type for ADVANCE
      if (Math.abs(diff) < 0.01) {
        settlementType = 'SETTLED';
        settlementAmount = 0;
      } else if (diff > 0) {
        settlementType = 'EXCESS';
        settlementAmount = diff; // Employee is owed this amount
      } else {
        settlementType = 'DEFICIT';
        settlementAmount = Math.abs(diff); // Employee needs to repay this amount
      }
    }
    
    let finalTicketStatus = CONFIG.TICKET_STATUS.VERIFIED;
    let finalClosedAt = '';
    let historyNote = verificationNotes || 'Bills verified and approved';
    
    if (ticket.expenseType === CONFIG.EXPENSE_TYPE.REIMBURSEMENT) {
      finalTicketStatus = CONFIG.TICKET_STATUS.REIMBURSED_CLOSED;
      finalClosedAt = now;
      historyNote = 'Bills verified. REIMBURSEMENT ticket auto-closed.' + 
                   (verificationNotes ? ' Note: ' + verificationNotes : '');
    } else if (ticket.expenseType === CONFIG.EXPENSE_TYPE.ADVANCE) {
      // ADVANCE tickets go to SETTLED status with settlement info
      finalTicketStatus = CONFIG.TICKET_STATUS.SETTLED;
      finalClosedAt = now;
      historyNote = 'Bills verified. Settlement: ' + settlementType + 
                   (settlementAmount > 0 ? ' - \u20b9' + settlementAmount.toLocaleString('en-IN') : '') +
                   (verificationNotes ? '. Note: ' + verificationNotes : '');
    }
    
    this.handler.updateByTicket(ticketNumber, {
      ticketStatus:         finalTicketStatus,
      closedAt:             finalClosedAt,
      billVerificationJSON: envelope,
      billVerifiedAt:       ticket.billVerifiedAt || now,
      totalBillAmount:      billTotalAmount,
      // Settlement fields
      settlementType:       settlementType,
      settlementAmount:     settlementAmount,
      settledAt:            ticket.expenseType === CONFIG.EXPENSE_TYPE.ADVANCE ? now : '',
      historyJSON:          JSON.parse(
                              this.sh.addHistoryEntry(ticket.historyJSON, 
                                finalTicketStatus === CONFIG.TICKET_STATUS.REIMBURSED_CLOSED 
                                  ? 'REIMBURSEMENT_CLOSED' 
                                  : finalTicketStatus === CONFIG.TICKET_STATUS.SETTLED
                                    ? 'SETTLED'
                                    : 'BILLS_VERIFIED',
                                { userId, name: user.name, role: ud.role },
                                historyNote)
                            )
    });

    return { 
      ticketStatus: finalTicketStatus, 
      settlementType: settlementType,
      settlementAmount: settlementAmount,
      message: ticket.expenseType === CONFIG.EXPENSE_TYPE.REIMBURSEMENT
        ? 'Bills approved and ticket closed'
        : ticket.expenseType === CONFIG.EXPENSE_TYPE.ADVANCE
          ? 'Bills approved. Settlement: ' + settlementType + (settlementAmount > 0 ? ' (\u20b9' + settlementAmount.toLocaleString('en-IN') + ')' : '')
          : 'Bills approved.'
    };
  }

  /**
   * requestBillCorrections — accounts requests corrections
   */
  requestBillCorrections({ token, ticketNumber, corrections }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    if (!this.sh.canVerifyBills(ud.role)) {
      throw new Error('Only Accounts or Admin can request bill corrections.');
    }

    Validator.requireNonEmpty(corrections, 'Correction notes');

    const ticket = this.handler.getByTicket(ticketNumber);
    if (!ticket) throw new Error('Ticket not found.');

    if (ticket.ticketStatus !== CONFIG.TICKET_STATUS.BILLED &&
        ticket.ticketStatus !== CONFIG.TICKET_STATUS.BILL_CORRECTION_NEEDED) {
      throw new Error('Ticket is not in BILLED state.');
    }

    const envelope = ticket.billVerificationJSON || [];
    envelope.push({
      action:      'CORRECTION_NEEDED',
      by:          userId,
      name:        user.name,
      timestamp:   new Date().toISOString(),
      notes:       corrections,
      corrections
    });

    this.handler.updateByTicket(ticketNumber, {
      ticketStatus:         CONFIG.TICKET_STATUS.BILL_CORRECTION_NEEDED,
      billVerificationJSON: envelope,
      historyJSON:          JSON.parse(
                              this.sh.addHistoryEntry(ticket.historyJSON, 'BILL_CORRECTION_NEEDED',
                                { userId, name: user.name, role: ud.role },
                                'Correction requested: ' + corrections)
                            )
    });

    try {
      _mailBillCorrectionNeeded(ticket, corrections);
    } catch(e) {
      Logger.log('Mail stub failed: ' + e.message);
    }
    return { ticketStatus: CONFIG.TICKET_STATUS.BILL_CORRECTION_NEEDED, message: 'Correction requested.' };
  }

  /**
   * rejectBills — accounts rejects bills
   */
  rejectBills({ token, ticketNumber, rejectionReason }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    if (!this.sh.canVerifyBills(ud.role)) {
      throw new Error('Only Accounts or Admin can reject bills.');
    }

    Validator.requireNonEmpty(rejectionReason, 'Rejection reason');

    const ticket = this.handler.getByTicket(ticketNumber);
    if (!ticket) throw new Error('Ticket not found.');

    const envelope = ticket.billVerificationJSON || [];
    envelope.push({
      action:    'REJECTED',
      by:        userId,
      name:      user.name,
      timestamp: new Date().toISOString(),
      notes:     rejectionReason
    });

    this.handler.updateByTicket(ticketNumber, {
      ticketStatus:         CONFIG.TICKET_STATUS.BILL_REJECTED,
      currentStage:         CONFIG.STAGE.COMPLETED,
      billVerificationJSON: envelope,
      closedAt:             new Date().toISOString(),
      historyJSON:          JSON.parse(
                              this.sh.addHistoryEntry(ticket.historyJSON, 'BILLS_REJECTED',
                                { userId, name: user.name, role: ud.role },
                                'Bills rejected: ' + rejectionReason)
                            )
    });

    try {
      _mailBillRejected(ticket, rejectionReason);
    } catch(e) {
      Logger.log('Mail stub failed: ' + e.message);
    }
    return { ticketStatus: CONFIG.TICKET_STATUS.BILL_REJECTED, message: 'Bills rejected. Ticket closed.' };
  }

  /**
   * closeTicket — accounts closes verified ticket with carry forward calculation
   */
  closeTicket({ token, ticketNumber, closureNotes }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    if (!this.sh.canProcessPayments(ud.role)) {
      throw new Error('Only Accounts or Admin can close tickets.');
    }

    const ticket = this.handler.getByTicket(ticketNumber);
    if (!ticket) throw new Error('Ticket not found.');

    if (ticket.ticketStatus !== CONFIG.TICKET_STATUS.VERIFIED && 
        ticket.ticketStatus !== CONFIG.TICKET_STATUS.SETTLED) {
      throw new Error('Ticket must be VERIFIED or SETTLED before closing.');
    }

    // Settlement info is already set during bill approval (approveBills)
    // For ADVANCE: settlementType, settlementAmount are already saved
    // For VERIFIED tickets (non-ADVANCE): just close directly
    const now = new Date().toISOString();
    const updates = {
      ticketStatus: CONFIG.TICKET_STATUS.CLOSED,
      currentStage: CONFIG.STAGE.COMPLETED,
      closedAt: now,
      historyJSON: JSON.parse(
        this.sh.addHistoryEntry(ticket.historyJSON, 'TICKET_CLOSED',
          { userId, name: user.name, role: ud.role },
          (closureNotes ? closureNotes + '. ' : '') + 'Ticket closed.')
      )
    };

    this.handler.updateByTicket(ticketNumber, updates);

    return {
      ticketStatus: CONFIG.TICKET_STATUS.CLOSED,
      message: 'Ticket closed.' + 
               (ticket.settlementType && ticket.settlementAmount > 0 
                 ? ' Settlement: ' + ticket.settlementType + ' ₹' + ticket.settlementAmount 
                 : '')
    };
  }

  /**
   * ✅ NEW: verifyBillItem — approve or reject individual bill items
   */
  verifyBillItem({ token, ticketNumber, itemIndex, action, notes }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    if (!this.sh.canProcessPayments(ud.role)) {
      throw new Error('Only Accounts or Admin can verify bills.');
    }

    const ticket = this.handler.getByTicket(ticketNumber);
    if (!ticket) throw new Error('Ticket not found.');

    if (ticket.ticketStatus !== CONFIG.TICKET_STATUS.BILLED) {
      throw new Error('Ticket must be in BILLED status for bill verification.');
    }

    // Get latest bill submission
    const billSubs = ticket.billJSON?.billSubmissions || [];
    if (!billSubs.length) throw new Error('No bills to verify.');

    const latestBill = billSubs[billSubs.length - 1];
    if (!Array.isArray(latestBill.items) || !latestBill.items[itemIndex]) {
      throw new Error('Bill item not found.');
    }

    // Initialize verification tracking if not exists
    if (!Array.isArray(ticket.billVerificationJSON)) {
      ticket.billVerificationJSON = [];
    }

    // Add verification record for this item
    const verification = {
      action: action.toUpperCase(), // APPROVED | REJECTED | HOLD | CORRECTION_REQUESTED
      billItemIndex: itemIndex,
      by: userId,
      name: user.name,
      timestamp: new Date().toISOString(),
      notes: notes || '',
      amount: latestBill.items[itemIndex].amount
    };

    ticket.billVerificationJSON.push(verification);

    this.handler.updateByTicket(ticketNumber, {
      billVerificationJSON: ticket.billVerificationJSON,
      historyJSON: JSON.parse(
        this.sh.addHistoryEntry(ticket.historyJSON, 'BILL_ITEM_VERIFIED',
          { userId, name: user.name, role: ud.role },
          'Bill item #' + (itemIndex + 1) + ' ' + action.toLowerCase() + '. Note: ' + (notes || '—'))
      )
    });

    return {
      message: 'Bill item verified: ' + action,
      verification
    };
  }

  /**
   * ✅ NEW: requestBillCorrections — request corrections on specific bills
   */
  requestBillCorrections({ token, ticketNumber, itemIndices, correctionNotes }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    if (!this.sh.canProcessPayments(ud.role)) {
      throw new Error('Only Accounts or Admin can request corrections.');
    }

    const ticket = this.handler.getByTicket(ticketNumber);
    if (!ticket) throw new Error('Ticket not found.');

    if (ticket.ticketStatus !== CONFIG.TICKET_STATUS.BILLED) {
      throw new Error('Ticket must be in BILLED status.');
    }

    const now = new Date().toISOString();
    this.handler.updateByTicket(ticketNumber, {
      ticketStatus: CONFIG.TICKET_STATUS.BILL_CORRECTION_NEEDED,
      remarksFromAccounts: correctionNotes || 'Corrections needed on submitted bills',
      historyJSON: JSON.parse(
        this.sh.addHistoryEntry(ticket.historyJSON, 'CORRECTION_REQUESTED',
          { userId, name: user.name, role: ud.role },
          'Corrections requested on items: ' + itemIndices.join(', ') + '. Note: ' + correctionNotes)
      )
    });

    return {
      ticketStatus: CONFIG.TICKET_STATUS.BILL_CORRECTION_NEEDED,
      message: 'Correction request sent to employee'
    };
  }

  /**
   * ✅ NEW: Helper to calculate total approved bill amount
   */
  _calculateApprovedBillTotal(ticket) {
    if (!ticket.billVerificationJSON || !Array.isArray(ticket.billVerificationJSON)) {
      return ticket.totalBillAmount || 0;
    }

    // Sum only APPROVED bills
    const approvedItems = new Set();
    ticket.billVerificationJSON.forEach(v => {
      if (v.action === 'APPROVED') {
        approvedItems.add(v.billItemIndex);
      }
    });

    const billSubs = ticket.billJSON?.billSubmissions || [];
    if (!billSubs.length) return ticket.totalBillAmount || 0;

    const latestBill = billSubs[billSubs.length - 1];
    if (!Array.isArray(latestBill.items)) return 0;

    return latestBill.items.reduce((sum, item, idx) => {
      return approvedItems.has(idx) ? sum + (item.amount || 0) : sum;
    }, 0);
  }

  /**
   * forceCloseTicket — admin override to close any ticket
   */
  forceCloseTicket({ token, ticketNumber, forceReason }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    if (!this.sh.isAdmin(ud.role)) {
      throw new Error('Only Admin can force-close tickets.');
    }

    Validator.requireNonEmpty(forceReason, 'Force closure reason');

    const ticket = this.handler.getByTicket(ticketNumber);
    if (!ticket) throw new Error('Ticket not found.');

    this.handler.updateByTicket(ticketNumber, {
      ticketStatus: CONFIG.TICKET_STATUS.FORCE_CLOSED,
      currentStage: CONFIG.STAGE.COMPLETED,
      closedAt:     new Date().toISOString(),
      historyJSON:  JSON.parse(
                      this.sh.addHistoryEntry(ticket.historyJSON, 'FORCE_CLOSED',
                        { userId, name: user.name, role: ud.role },
                        'Force closed by admin. Reason: ' + forceReason)
                    )
    });

    return { ticketStatus: CONFIG.TICKET_STATUS.FORCE_CLOSED, message: 'Ticket force-closed.' };
  }

  /**
   * adminWaiveBill — admin waives bill requirement
   */
  adminWaiveBill({ token, ticketNumber, notes }) {
    const ud     = this._auth(token);
    const userId = ud.userId || ud.email;
    const user   = this.sh.getUserByEmail(ud.email);

    if (!this.sh.isAdmin(ud.role)) {
      throw new Error('Only Admin can waive bill requirement.');
    }

    Validator.requireNonEmpty(notes, 'Waive notes');

    const ticket = this.handler.getByTicket(ticketNumber);
    if (!ticket) throw new Error('Ticket not found.');

    this.handler.updateByTicket(ticketNumber, {
      ticketStatus: CONFIG.TICKET_STATUS.ADMIN_WAIVED,
      currentStage: CONFIG.STAGE.COMPLETED,
      closedAt:     new Date().toISOString(),
      historyJSON:  JSON.parse(
                      this.sh.addHistoryEntry(ticket.historyJSON, 'ADMIN_WAIVED',
                        { userId, name: user.name, role: ud.role },
                        'Bill requirement waived by admin. Notes: ' + notes)
                    )
    });

    return { ticketStatus: CONFIG.TICKET_STATUS.ADMIN_WAIVED, message: 'Bill requirement waived.' };
  }

  // ════════════════════════════════════════════════════════════
  //  ACCOUNTS DASHBOARD & RECONCILIATION
  // ════════════════════════════════════════════════════════════

  /**
   * getAccountsDashboard
   */
  getAccountsDashboard({ token }) {
    const ud = this._auth(token);
    if (!this.sh.canProcessPayments(ud.role)) {
      throw new Error('Only Accounts or Admin can access dashboard.');
    }

    const rows = this.handler.getAll().filter(r => r.ticketNumber && r.ticketNumber !== '');
    const S    = CONFIG.TICKET_STATUS;

    let pendingPayment = 0, pendingBills = 0, pendingVerification = 0,
        totalClosed = 0, pendingAmount = 0, paidAmount = 0;

    rows.forEach(t => {
      const status = t.ticketStatus;
      if (status === S.CREATED || status === S.PAYMENT_PARTIAL || status === S.PAYMENT_HOLD) {
        pendingPayment++;
        pendingAmount += parseFloat(t.totalApprovedAmount) || 0;
      }
      if (status === S.PAYMENT_FULL) pendingBills++;
      if (status === S.BILLED || status === S.BILL_CORRECTION_NEEDED) pendingVerification++;
      if (CONFIG.TERMINAL_TICKET_STATUSES.includes(status) || status === S.VERIFIED) totalClosed++;
      paidAmount += parseFloat(t.totalPaidAmount) || 0;
    });

    return {
      pendingPayment, pendingBills, pendingVerification,
      totalClosed, pendingAmount, paidAmount,
      total: rows.length
    };
  }

  /**
   * getPaymentReconciliation
   */
  getPaymentReconciliation({ token }) {
    const ud = this._auth(token);
    if (!this.sh.canProcessPayments(ud.role)) {
      throw new Error('Only Accounts or Admin can view reconciliation.');
    }

    const rows = this.handler.getAll().filter(r => r.ticketNumber && r.ticketNumber !== '');
    let totalApproved = 0, totalPaid = 0;

    const details = rows.map(t => {
      const approved = parseFloat(t.totalApprovedAmount) || 0;
      const paid     = parseFloat(t.totalPaidAmount)     || 0;
      totalApproved += approved;
      totalPaid     += paid;
      return {
        ticketNumber:      t.ticketNumber,
        referenceNumber:   t.referenceNumber,
        employeeName:      t.name,
        department:        t.department,
        transactionType:   t.transactionType,
        approvedAmount:    approved,
        paidAmount:        paid,
        outstanding:       approved - paid,
        ticketStatus:      t.ticketStatus
      };
    });

    return {
      summary: { totalApproved, totalPaid, outstanding: totalApproved - totalPaid },
      details
    };
  }

  /**
   * getTicketAuditTrail
   */
  getTicketAuditTrail({ token, ticketNumber }) {
    const ud     = this._auth(token);
    const ticket = this.handler.getByTicket(ticketNumber);
    if (!ticket) throw new Error('Ticket not found.');

    const userId  = ud.userId || ud.email;
    const isOwner = ticket.userId === userId;
    if (!isOwner && ud.role !== CONFIG.ROLES.ACCOUNTS && !this.sh.isAdmin(ud.role)) {
      throw new Error('Only the ticket owner, Accounts, or Admin can view audit trails.');
    }

    return {
      ticketNumber:        ticket.ticketNumber,
      referenceNumber:     ticket.referenceNumber,
      employeeName:        ticket.name,
      ticketStatus:        ticket.ticketStatus,
      transactionType:     ticket.transactionType,
      totalApprovedAmount: ticket.totalApprovedAmount,
      totalPaymentAmount:  ticket.totalPaidAmount,
      workflowHistory:     ticket.historyJSON           || [],
      paymentHistory:      ticket.paymentJSON?.payments  || [],
      billSubmissions:     ticket.billJSON?.billSubmissions || [],
      billVerifications:   ticket.billVerificationJSON   || []
    };
  }
}

// ── INTERNAL HELPERS ──────────────────────────────────────────
function _sumItems(items) {
  if (!Array.isArray(items)) return 0;
  return items.reduce((s, i) => s + (parseFloat(i.amount) || 0), 0);
}

function _latestBillTotal(billEnvelope) {
  try {
    const subs = billEnvelope.billSubmissions;
    if (!Array.isArray(subs) || subs.length === 0) return 0;
    return parseFloat(subs[subs.length - 1].totalBillAmount) || 0;
  } catch (e) { return 0; }
}

function _sumPayments(paymentEnvelope) {
  try {
    const pmts = paymentEnvelope.payments;
    if (!Array.isArray(pmts)) return 0;
    return pmts.reduce((s, p) => s + (parseFloat(p.amount) || 0), 0);
  } catch (e) { return 0; }
}

// ── MAIL STUBS ────────────────────────────────────────────────
function _mailExpenseRejected(expense, reason) {
  // TODO: implement
}


function _mailBillCorrectionNeeded(ticket, corrections) {
  // TODO: implement
}

function _mailBillRejected(ticket, reason) {
  // TODO: implement
}

function _mailPaymentDelayed(ticket, notes) {
  // TODO: implement
}

// ── SINGLETON ──────────────────────────────────────────────────
const _service = new UnifiedService();

// ── EXPOSED RPC ────────────────────────────────────────────────
// EXPENSES
function createExpense(p)                { return _wrap(q => _service.createExpense(q))(p); }
function getUserExpenses(p)              { return _wrap(q => _service.getUserExpenses(q))(p); }
function getExpense(p)                   { return _wrap(q => _service.getExpense(q))(p); }
function forwardExpenseToOpHead(p)       { return _wrap(q => _service.forwardExpenseToOpHead(q))(p); }
function approveExpense(p)               { return _wrap(q => _service.approveExpense(q))(p); }
function rejectExpense(p)                { return _wrap(q => _service.rejectExpense(q))(p); }
function cancelExpense(p)                { return _wrap(q => _service.cancelExpense(q))(p); }
function holdExpense(p)                  { return _wrap(q => _service.holdExpense(q))(p); }
function releaseExpenseHold(p)           { return _wrap(q => _service.releaseExpenseHold(q))(p); }
function resubmitExpense(p)              { return _wrap(q => _service.resubmitExpense(q))(p); }
function updateExpenseItems(p)           { return _wrap(q => _service.updateExpenseItems(q))(p); }
function updateExpenseBankDetails(p)     { return _wrap(q => _service.updateBankDetails(q))(p); }
function submitBill(p)                   { return _wrap(q => _service.submitBill(q))(p); }
function getExpenseStats(p)              { return _wrap(q => _service.getExpenseStats(q))(p); }
function getPendingExpenseApprovals(p)   { return _wrap(q => _service.getPendingExpenseApprovals(q))(p); }
function processExpensePayment(p)     { return _wrap(q => _service.processExpensePayment(q))(p); }

// TICKETS / PAYMENTS / BILLS
function getTickets(p)                 { return _wrap(q => _service.getTickets(q))(p); }
function getTicketDetail(p)            { return _wrap(q => _service.getTicketDetail(q))(p); }
function getPendingPayments(p)         { return _wrap(q => _service.getPendingPayments(q))(p); }
function getPendingBillVerification(p) { return _wrap(q => _service.getPendingBillVerification(q))(p); }
function processPayment(p)             { return _wrap(q => _service.processPayment(q))(p); }
function flagPaymentDelay(p)           { return _wrap(q => _service.flagPaymentDelay(q))(p); }
function submitBills(p)                { return _wrap(q => _service.submitBillsOnTicket(q))(p); }
function approveBills(p)               { return _wrap(q => _service.approveBills(q))(p); }
function requestBillCorrections(p)     { return _wrap(q => _service.requestBillCorrections(q))(p); }
function verifyBillItem(p)             { return _wrap(q => _service.verifyBillItem(q))(p); }
function rejectBills(p)                { return _wrap(q => _service.rejectBills(q))(p); }
function closeTicket(p)                { return _wrap(q => _service.closeTicket(q))(p); }
function forceCloseTicket(p)           { return _wrap(q => _service.forceCloseTicket(q))(p); }
function adminWaiveBill(p)             { return _wrap(q => _service.adminWaiveBill(q))(p); }

// ACCOUNTS DASHBOARD
function getAccountsDashboard(p)       { return _wrap(q => _service.getAccountsDashboard(q))(p); }
function getPaymentReconciliation(p)   { return _wrap(q => _service.getPaymentReconciliation(q))(p); }
function getTicketAuditTrail(p)        { return _wrap(q => _service.getTicketAuditTrail(q))(p); }

// PAYMENT HOLD
function holdPayment(p)                { return _wrap(q => _service.holdPayment(q))(p); }
function releasePaymentHold(p)         { return _wrap(q => _service.releasePaymentHold(q))(p); }