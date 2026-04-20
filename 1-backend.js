/**
 * ============================================================
 *  1-backend.js  —  App Entry Points & Auth
 *  Expense Management System (GAS)
 *
 *  Responsibilities:
 *    • doGet()     → serves the SPA shell (index.html)
 *    • include_()  → partial-template helper
 *    • Auth RPC:   login, logout, getUsers
 *    • initializeApp() → one-time sheet setup
 *
 *  Depends on: 0-shared_utils.js
 * ============================================================
 */

// ── TOKEN SCHEMA VERSION ──────────────────────────────────────
// Bump this when the cached token structure changes, to auto-invalidate old sessions.
var TOKEN_SCHEMA_VERSION = 3;

// ── APP CLASS ─────────────────────────────────────────────────
class App {
  constructor() {
    this.db    = SpreadsheetApp.openById(CONFIG.SettingsID);
    this.cache = CacheService.getScriptCache();
    this.sh    = new SharedHelper(this.db, this.cache);
  }

  // ── Token management ────────────────────────────────────────
  createToken(user) {
    const token     = Utilities.getUuid();
    const tokenData = {
      _v:         TOKEN_SCHEMA_VERSION,
      userId:     user.UserID || user.email,
      email:      user.email,
      role:       user.role,
      name:       user.name,
      jobTitle:   user.jobTitle    || '',
      department: user.department  || '',
      project:    user.project     || '',
      district:   user.district    || '',
      managerName: user.managerName || '',
      isActive:   user.isActive
    };
    this.cache.put(token, JSON.stringify(tokenData), CONFIG.CACHE_TTL);
    return token;
  }

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
    // Sliding TTL
    this.cache.put(token, raw, CONFIG.CACHE_TTL);
    return parsed;
  }

  // ── Auth ────────────────────────────────────────────────────
  login(params) {
    // Token-based re-auth (session restore)
    if (params.token) {
      const ud = this.validateToken(params.token);
      if (!ud) throw new Error('Session expired. Please log in again.');
      const user = this._getUserByEmail(ud.email);
      if (!user) throw new Error('User not found.');
      const safe = Object.assign({}, user);
      delete safe.password;
      return { token: params.token, user: safe };
    }

    // Credential-based login
    const email    = (params.email    || '').toString().trim().toLowerCase();
    const password = (params.password || '').toString();
    if (!email || !password) throw new Error('Email and password are required.');

    const user = this._getUserByEmail(email);
    if (!user) throw new Error('Invalid email or password.');
    if (user.password !== password) throw new Error('Invalid email or password.');
    if (!Validator.parseBoolean(user.isActive)) throw new Error('Account is inactive.');

    const token   = this.createToken(user);
    const safe    = Object.assign({}, user);
    delete safe.password;
    return { token, user: safe };
  }

  logout(params) {
    if (params.token) this.cache.remove(params.token);
    return { ok: true };
  }

  // ── User data ────────────────────────────────────────────────
  /**
   * Returns all active users (password excluded).
   * Any authenticated user can call this (needed for manager lookups, etc.)
   */
  getUsers(params) {
    const ud = this.validateToken(params.token);
    if (!ud) throw new Error('Unauthorized.');
    // Delegate — no more manual sheet scan here
    return this.sh.getAllUsers()
      .filter(u => Validator.parseBoolean(u.isActive));
  }

  // ── Private helpers ─────────────────────────────────────────
  _getUserByEmail(email) {
    const ws   = this.db.getSheetByName(CONFIG.SHEET_NAME.USERS);
    const data = ws.getDataRange().getValues();
    const keys = data[0];
    const rows = data.slice(1);
    const idx  = keys.indexOf('email');

    const row = rows.find(r => (r[idx] || '').toString().trim().toLowerCase() === email.toLowerCase());
    if (!row) return null;

    const obj = {};
    keys.forEach((k, i) => { obj[k] = row[i]; });
    return obj;
  }
}

// ── SINGLETON ─────────────────────────────────────────────────
var _app = new App();

// ── GAS ENTRY POINTS ───────────────────────────────────────────
function doGet() {
  return HtmlService
    .createTemplateFromFile(CONFIG.INDEX)
    .evaluate()
    .setTitle(CONFIG.NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include_(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

// ── AUTH RPC ──────────────────────────────────────────────────
function login(p)    { return _wrap(q => _app.login(q))(p); }
function logout(p)   { return _wrap(q => _app.logout(q))(p); }
function getUsers(p) { return _wrap(q => _app.getUsers(q))(p); }

// ── GLOBAL TOKEN VALIDATOR ───────────────────────────────────
/**
 * Used by service files that need to validate a token without
 * instantiating a full App object.
 *
 * Returns the parsed token payload or null.
 */
function validateTokenAndGetUser(token) {
  return _app.validateToken(token);
}

// ── ONE-TIME SETUP ────────────────────────────────────────────
/**
 * Run ONCE from the GAS editor to create all required sheets.
 * Safe to re-run — existing sheets are left untouched.
 */
function initializeApp() {
  const db = SpreadsheetApp.openById(CONFIG.SettingsID);

  _ensureSheet(db, CONFIG.SHEET_NAME.USERS, CONFIG.SHEET_HEADERS.USERS, {
    columnWidths: { 1: 160, 3: 220, 4: 220 }
  });

  _ensureSheet(db, CONFIG.SHEET_NAME.EXPENSES, CONFIG.SHEET_HEADERS.EXPENSES, {
    columnWidths: { 1: 170, 2: 140, 5: 210 },
    freezeRows: 1
  });

  _getOverflowManager().ensureSheet();

  Logger.log('initializeApp complete — all sheets ready');
}

function _ensureSheet(db, name, headers, opts) {
  opts = opts || {};
  let sheet = db.getSheetByName(name);
  if (sheet) { Logger.log('Sheet "' + name + '" already exists — skipped.'); return sheet; }

  sheet = db.insertSheet(name);
  sheet.appendRow(headers);

  const hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setFontWeight('bold');
  hRange.setBackground('#4A5568');
  hRange.setFontColor('#FFFFFF');
  hRange.setWrap(false);

  if (opts.columnWidths) {
    Object.keys(opts.columnWidths).forEach(col => {
      sheet.setColumnWidth(Number(col), opts.columnWidths[col]);
    });
  }
  if (opts.freezeRows) sheet.setFrozenRows(opts.freezeRows);

  Logger.log('Created sheet "' + name + '" with ' + headers.length + ' columns.');
  return sheet;
}