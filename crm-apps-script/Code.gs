/**
 * CRM Console — Customer balances & collections (Dynamics 365 data)
 * Google Apps Script backend. Google Sheets is the database.
 *
 * First run: open the Apps Script editor and run `setup()` once.
 * Default login: admin / admin123  (change it from the Users page).
 */

// ───────────────────────────── Schema ─────────────────────────────
// Each field: [key, sheet header]. Headers match the Dynamics 365 export
// names so exported data can be pasted directly under them.
const SCHEMA = {
  Customers: [
    ['account', 'Customer account'], ['name', 'Name'], ['group', 'Customer group'],
    ['currency', 'Currency'], ['contactPerson', 'Contact Person'], ['phone', 'Contact Number'],
    ['email', 'Email'], ['address', 'Address'], ['city', 'City'], ['country', 'Country'],
    ['taxNo', 'Tax Number'], ['creditLimit', 'Credit Limit'], ['termsDays', 'Payment Terms Days'],
    ['collector', 'Collector'], ['status', 'Status'], ['notes', 'Notes'],
    ['createdAt', 'Created date and time'], ['updatedBy', 'Updated By'], ['updatedAt', 'Updated At']
  ],
  Transactions: [
    ['voucher', 'Voucher'], ['type', 'Transaction type'], ['account', 'Customer account'],
    ['date', 'Date'], ['description', 'Description'], ['dueDate', 'Due Date'],
    ['amountCur', 'Amount in transaction currency'], ['currency', 'Currency'],
    ['exchange', 'Exchange rate'], ['creditMST', 'Credit in local currency'],
    ['debitMST', 'Debit in local currency'], ['userAdd', 'user add'],
    ['createdById', 'created_by_user_id'], ['createdByName', 'created_by_name'],
    ['createdAt', 'created_at'], ['updatedByName', 'updated_by_name'], ['updatedAt', 'updated_at'],
    ['isDeleted', 'is_deleted'], ['postStatus', 'posting_status'], ['postedBy', 'posted_by'],
    ['postedAt', 'posted_at'], ['rowId', 'row_id']
  ],
  CustomerGroups: [['groupId', 'GroupId'], ['groupName', 'GroupName']],
  Users: [
    ['userId', 'UserId'], ['username', 'Username'], ['fullName', 'Full Name'], ['email', 'Email'],
    ['passwordHash', 'PasswordHash'], ['salt', 'Salt'], ['role', 'Role'], ['pages', 'Pages'],
    ['canAdd', 'CanAdd'], ['canEdit', 'CanEdit'], ['canDelete', 'CanDelete'], ['canPost', 'CanPost'],
    ['isCollector', 'IsCollector'], ['scope', 'DataScope'], ['active', 'Active'],
    ['lastLogin', 'LastLogin'], ['createdAt', 'CreatedAt']
  ],
  FollowUps: [
    ['id', 'FollowUpId'], ['account', 'Customer account'], ['voucher', 'Voucher'],
    ['actionDate', 'Action Date'], ['actionType', 'Action Type'], ['note', 'Note'],
    ['promiseDate', 'Promise Date'], ['promiseAmount', 'Promise Amount'], ['status', 'Status'],
    ['createdBy', 'Created By'], ['createdAt', 'Created At']
  ],
  Settings: [['key', 'Key'], ['value', 'Value']]
};

const ID_FIELD = { Customers: 'account', Transactions: 'rowId', CustomerGroups: 'groupId', Users: 'userId', FollowUps: 'id' };
const DATE_FIELDS = ['date', 'dueDate', 'createdAt', 'updatedAt', 'postedAt', 'lastLogin', 'actionDate', 'promiseDate'];
const NUMBER_FIELDS = ['amountCur', 'exchange', 'creditMST', 'debitMST', 'creditLimit', 'termsDays', 'promiseAmount'];
const BOOL_FIELDS = ['isDeleted', 'canAdd', 'canEdit', 'canDelete', 'canPost', 'isCollector', 'active'];

const ALL_PAGES = ['dashboard', 'customers', 'transactions', 'groups', 'balances', 'paid',
  'overdue', 'followup', 'collection', 'import', 'users'];

// Which entity each page edits (used for permission checks on writes).
const ENTITY_PAGE = { Customers: 'customers', Transactions: 'transactions', CustomerGroups: 'groups', FollowUps: 'followup', Users: 'users' };

const SESSION_HOURS = 8;
const TZ = Session.getScriptTimeZone();

// ───────────────────────────── Web app ─────────────────────────────
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('CRM Console')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('CRM')
    .addItem('Setup / repair sheets', 'setup')
    .addItem('Load demo data', 'loadDemoData')
    .addToUi();
}

// ───────────────────────────── Setup ─────────────────────────────
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(SCHEMA).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = SCHEMA[name].map(function (f) { return f[1]; });
    const lastCol = Math.max(sh.getLastColumn(), 1);
    const existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    const missing = headers.filter(function (h) { return existing.indexOf(h) === -1; });
    if (sh.getLastRow() === 0 || existing.join('') === '') {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else if (missing.length) {
      sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
    }
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#111114').setFontColor('#fafafa');
  });
  // Default admin
  if (!readTable('Users').some(function (u) { return String(u.username).toLowerCase() === 'admin'; })) {
    const salt = Utilities.getUuid();
    appendRows('Users', [{
      userId: Utilities.getUuid(), username: 'admin', fullName: 'Administrator', email: '',
      passwordHash: hash_('admin123', salt), salt: salt, role: 'Admin', pages: ALL_PAGES.join(','),
      canAdd: true, canEdit: true, canDelete: true, canPost: true, isCollector: false,
      scope: 'all', active: true, createdAt: new Date()
    }]);
  }
  const settings = readTable('Settings');
  const defaults = { AppName: 'CRM Console', LocalCurrency: 'SAR', CompanyName: 'My Company' };
  const toAdd = Object.keys(defaults).filter(function (k) {
    return !settings.some(function (s) { return s.key === k; });
  }).map(function (k) { return { key: k, value: defaults[k] }; });
  if (toAdd.length) appendRows('Settings', toAdd);
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
  return 'Setup complete';
}

// ───────────────────────────── Sheet helpers ─────────────────────────────
function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" not found. Run setup() first.');
  return sh;
}

function headerMap_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  const map = {};
  headers.forEach(function (h, i) { map[h] = i; });
  return { headers: headers, map: map };
}

function fieldIndex_(name, hm) {
  const out = {};
  SCHEMA[name].forEach(function (f) { if (hm.map[f[1]] !== undefined) out[f[0]] = hm.map[f[1]]; });
  return out;
}

/** Reads a sheet into plain objects (dates → yyyy-MM-dd strings for transport). */
function readTable(name, keepDates) {
  const sh = sheet_(name);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const hm = headerMap_(sh);
  const idx = fieldIndex_(name, hm);
  const values = sh.getRange(2, 1, lastRow - 1, hm.headers.length).getValues();
  const out = [];
  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    if (row.join('') === '') continue;
    const o = { _row: r + 2 };
    for (const k in idx) o[k] = normalizeValue_(k, row[idx[k]], keepDates);
    out.push(o);
  }
  return out;
}

function normalizeValue_(key, v, keepDates) {
  if (DATE_FIELDS.indexOf(key) > -1) {
    const d = toDate_(v);
    if (!d) return '';
    if (keepDates) return d;
    return Utilities.formatDate(d, TZ, (key === 'createdAt' || key === 'updatedAt' || key === 'postedAt' || key === 'lastLogin') ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
  }
  if (NUMBER_FIELDS.indexOf(key) > -1) return toNum_(v);
  if (BOOL_FIELDS.indexOf(key) > -1) return toBool_(v);
  return v === null || v === undefined ? '' : String(v).trim();
}

function toDate_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v) ? null : v;
  if (typeof v === 'number') return new Date(Math.round((v - 25569) * 86400000)); // Excel serial
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0));
  m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?)?/i);
  if (m) { // D365 exports usually use M/D/YYYY
    let h = +(m[4] || 0);
    if (m[6] && /pm/i.test(m[6]) && h < 12) h += 12;
    if (m[6] && /am/i.test(m[6]) && h === 12) h = 0;
    const a = +m[1], b = +m[2];
    const month = a > 12 ? b : a, day = a > 12 ? a : b;
    return new Date(+m[3], month - 1, day, h, +(m[5] || 0));
  }
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function toNum_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/[^\d.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function toBool_(v) {
  if (typeof v === 'boolean') return v;
  return /^(true|yes|1|y|نعم)$/i.test(String(v).trim());
}

function toSheetValue_(key, v) {
  if (DATE_FIELDS.indexOf(key) > -1) return toDate_(v) || '';
  if (NUMBER_FIELDS.indexOf(key) > -1) return v === '' || v === null || v === undefined ? '' : toNum_(v);
  if (BOOL_FIELDS.indexOf(key) > -1) return toBool_(v);
  return v === null || v === undefined ? '' : v;
}

function appendRows(name, objs) {
  if (!objs.length) return;
  const sh = sheet_(name);
  const hm = headerMap_(sh);
  const idx = fieldIndex_(name, hm);
  const rows = objs.map(function (o) {
    const row = new Array(hm.headers.length).fill('');
    for (const k in idx) if (o[k] !== undefined) row[idx[k]] = toSheetValue_(k, o[k]);
    return row;
  });
  const chunk = 5000;
  for (let i = 0; i < rows.length; i += chunk) {
    const part = rows.slice(i, i + chunk);
    sh.getRange(sh.getLastRow() + 1, 1, part.length, hm.headers.length).setValues(part);
  }
}

function findRow_(name, id) {
  const sh = sheet_(name);
  const hm = headerMap_(sh);
  const col = fieldIndex_(name, hm)[ID_FIELD[name]];
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, col + 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === String(id).trim()) return i + 2;
  return -1;
}

function updateRow_(name, rowNum, patch) {
  const sh = sheet_(name);
  const hm = headerMap_(sh);
  const idx = fieldIndex_(name, hm);
  const range = sh.getRange(rowNum, 1, 1, hm.headers.length);
  const row = range.getValues()[0];
  for (const k in patch) if (idx[k] !== undefined) row[idx[k]] = toSheetValue_(k, patch[k]);
  range.setValues([row]);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function getSettings_() {
  const o = {};
  readTable('Settings').forEach(function (s) { o[s.key] = s.value; });
  return o;
}

// ───────────────────────────── Auth ─────────────────────────────
function hash_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + password, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function login(username, password) {
  const user = readTable('Users').filter(function (u) {
    return String(u.username).toLowerCase() === String(username || '').trim().toLowerCase();
  })[0];
  if (!user || !user.active || hash_(password, user.salt) !== user.passwordHash) {
    throw new Error('INVALID_LOGIN');
  }
  user.isAdmin = user.role === 'Admin';
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('sess_' + token, user.userId, SESSION_HOURS * 3600);
  updateRow_('Users', user._row, { lastLogin: new Date() });
  return { token: token, user: publicUser_(user), settings: getSettings_(), lookups: lookups_(user) };
}

function logout(token) {
  CacheService.getScriptCache().remove('sess_' + token);
  return true;
}

function resume(token) {
  const user = auth_(token);
  return { token: token, user: publicUser_(user), settings: getSettings_(), lookups: lookups_(user) };
}

function auth_(token, page) {
  const userId = token && CacheService.getScriptCache().get('sess_' + token);
  if (!userId) throw new Error('SESSION_EXPIRED');
  const user = readTable('Users').filter(function (u) { return u.userId === userId; })[0];
  if (!user || !user.active) throw new Error('SESSION_EXPIRED');
  CacheService.getScriptCache().put('sess_' + token, userId, SESSION_HOURS * 3600); // sliding
  user.isAdmin = user.role === 'Admin';
  user.pageList = user.isAdmin ? ALL_PAGES.slice() : String(user.pages || '').split(',').map(function (p) { return p.trim(); }).filter(String);
  if (page && !user.isAdmin && user.pageList.indexOf(page) === -1) throw new Error('NO_PERMISSION');
  return user;
}

function requireAction_(user, action) {
  if (user.isAdmin) return;
  const flag = { add: 'canAdd', edit: 'canEdit', delete: 'canDelete', post: 'canPost' }[action];
  if (!user[flag]) throw new Error('NO_PERMISSION');
}

function publicUser_(u) {
  const isAdmin = u.role === 'Admin';
  return {
    userId: u.userId, username: u.username, fullName: u.fullName, email: u.email, role: u.role,
    pages: isAdmin ? ALL_PAGES : String(u.pages || '').split(',').map(function (p) { return p.trim(); }).filter(String),
    canAdd: isAdmin || u.canAdd, canEdit: isAdmin || u.canEdit, canDelete: isAdmin || u.canDelete,
    canPost: isAdmin || u.canPost, isCollector: u.isCollector, scope: u.scope || 'all', isAdmin: isAdmin
  };
}

/** Customers the user is allowed to see (DataScope = own → only their collector accounts). */
function visibleCustomers_(user) {
  const all = readTable('Customers');
  if (user.isAdmin || user.scope !== 'own') return all;
  return all.filter(function (c) { return c.collector === user.username; });
}

function lookups_(user) {
  const users = readTable('Users');
  return {
    groups: readTable('CustomerGroups').map(function (g) { return { id: g.groupId, name: g.groupName }; }),
    collectors: users.filter(function (u) { return u.isCollector && u.active; })
      .map(function (u) { return { username: u.username, name: u.fullName || u.username }; }),
    customers: visibleCustomers_(user).map(function (c) { return { account: c.account, name: c.name }; }),
    transactionTypes: ['Invoice', 'Payment', 'Credit note', 'Debit note', 'Settlement', 'Opening balance', 'Adjustment']
  };
}

function getLookups(token) { return lookups_(auth_(token)); }

// ───────────────────────────── Generic list ─────────────────────────────
/**
 * Server-side search / filter / sort / paging so big D365 tables stay fast on mobile.
 * opts: { q, filters: {field: value}, dateFrom, dateTo, sort, desc, page, pageSize }
 */
function listRecords(token, entity, opts) {
  opts = opts || {};
  const page = { Customers: 'customers', Transactions: 'transactions', CustomerGroups: 'groups', Users: 'users', FollowUps: 'followup' }[entity];
  if (!page) throw new Error('Unknown entity');
  const user = auth_(token, page);
  if (entity === 'Users' && !user.isAdmin) throw new Error('NO_PERMISSION');
  if (entity === 'Transactions') ensureRowIds_();
  let rows;
  if (entity === 'Customers') rows = visibleCustomers_(user);
  else rows = readTable(entity);

  if (entity === 'Transactions' || entity === 'FollowUps') {
    const allowed = {};
    visibleCustomers_(user).forEach(function (c) { allowed[c.account] = c.name; });
    rows = rows.filter(function (r) { return allowed[r.account] !== undefined; });
    rows.forEach(function (r) { r.customerName = allowed[r.account]; });
    if (entity === 'Transactions') {
      rows.forEach(function (r) { if (!r.postStatus) r.postStatus = 'Posted'; });
      if (!opts.showDeleted) rows = rows.filter(function (r) { return !r.isDeleted; });
    }
  }
  if (entity === 'Users') rows = rows.map(function (u) { const o = Object.assign({}, u); delete o.passwordHash; delete o.salt; return o; });

  const q = String(opts.q || '').trim().toLowerCase();
  if (q) rows = rows.filter(function (r) {
    for (const k in r) if (k !== '_row' && String(r[k]).toLowerCase().indexOf(q) > -1) return true;
    return false;
  });
  const filters = opts.filters || {};
  Object.keys(filters).forEach(function (k) {
    const v = filters[k];
    if (v === '' || v === null || v === undefined) return;
    rows = rows.filter(function (r) { return String(r[k]) === String(v); });
  });
  const dateKey = entity === 'FollowUps' ? 'actionDate' : 'date';
  if (opts.dateFrom) rows = rows.filter(function (r) { return r[dateKey] && r[dateKey] >= opts.dateFrom; });
  if (opts.dateTo) rows = rows.filter(function (r) { return r[dateKey] && r[dateKey] <= opts.dateTo; });

  const sortKey = opts.sort || (entity === 'Transactions' ? 'date' : entity === 'FollowUps' ? 'actionDate' : ID_FIELD[entity]);
  const desc = opts.desc !== undefined ? opts.desc : (entity === 'Transactions' || entity === 'FollowUps');
  rows.sort(function (a, b) {
    const x = a[sortKey], y = b[sortKey];
    const c = (typeof x === 'number' && typeof y === 'number') ? x - y : String(x).localeCompare(String(y), undefined, { numeric: true });
    return desc ? -c : c;
  });

  const totals = {};
  if (entity === 'Transactions') {
    totals.debit = sum_(rows, 'debitMST');
    totals.credit = sum_(rows, 'creditMST');
  }
  const size = Math.min(Math.max(+opts.pageSize || 25, 5), 500);
  const total = rows.length;
  const p = Math.max(1, Math.min(+opts.page || 1, Math.ceil(total / size) || 1));
  return { rows: opts.all ? rows : rows.slice((p - 1) * size, p * size), total: total, page: p, pageSize: size, totals: totals };
}

function sum_(rows, key) {
  return round2_(rows.reduce(function (s, r) { return s + (+r[key] || 0); }, 0));
}
function round2_(n) { return Math.round(n * 100) / 100; }

// ───────────────────────────── Save / delete / post ─────────────────────────────
function saveRecord(token, entity, record, isNew) {
  const page = ENTITY_PAGE[entity];
  if (!page) throw new Error('Unknown entity');
  const user = auth_(token, page);
  requireAction_(user, isNew ? 'add' : 'edit');
  return withLock_(function () {
    const now = new Date();
    const who = user.fullName || user.username;
    const rec = Object.assign({}, record);
    delete rec._row;

    if (entity === 'Customers') {
      if (!rec.account || !rec.name) throw new Error('REQUIRED_FIELDS');
      if (user.scope === 'own' && !user.isAdmin) rec.collector = user.username;
      rec.updatedBy = who; rec.updatedAt = now;
      if (isNew) {
        if (findRow_('Customers', rec.account) > -1) throw new Error('DUPLICATE_ID');
        rec.createdAt = rec.createdAt || now;
        rec.status = rec.status || 'Active';
      }
    } else if (entity === 'Transactions') {
      if (!rec.account || !rec.date) throw new Error('REQUIRED_FIELDS');
      assertCustomerVisible_(user, rec.account);
      const debit = toNum_(rec.debitMST), credit = toNum_(rec.creditMST);
      rec.amountCur = rec.amountCur !== '' && rec.amountCur !== undefined ? rec.amountCur : round2_(credit - debit);
      rec.exchange = rec.exchange || 1;
      if (isNew) {
        rec.rowId = Utilities.getUuid();
        rec.createdById = user.userId; rec.createdByName = who; rec.createdAt = now;
        rec.userAdd = rec.userAdd || user.username;
        rec.isDeleted = false;
        rec.postStatus = 'Draft';
      } else {
        const existing = getById_(entity, rec.rowId);
        if (existing.postStatus === 'Posted') throw new Error('POSTED_LOCKED');
        ['createdById', 'createdByName', 'createdAt', 'userAdd', 'postStatus', 'postedBy', 'postedAt', 'isDeleted']
          .forEach(function (k) { delete rec[k]; });
      }
      rec.updatedByName = who; rec.updatedAt = now;
    } else if (entity === 'CustomerGroups') {
      if (!rec.groupId || !rec.groupName) throw new Error('REQUIRED_FIELDS');
      if (isNew && findRow_(entity, rec.groupId) > -1) throw new Error('DUPLICATE_ID');
    } else if (entity === 'FollowUps') {
      if (!rec.account) throw new Error('REQUIRED_FIELDS');
      assertCustomerVisible_(user, rec.account);
      if (isNew) { rec.id = Utilities.getUuid(); rec.createdBy = who; rec.createdAt = now; }
      rec.actionDate = rec.actionDate || now;
      rec.status = rec.status || 'Open';
    } else if (entity === 'Users') {
      if (!user.isAdmin) throw new Error('NO_PERMISSION');
      return saveUser_(rec, isNew);
    }

    if (isNew) appendRows(entity, [rec]);
    else {
      const rowNum = findRow_(entity, rec[ID_FIELD[entity]]);
      if (rowNum < 0) throw new Error('NOT_FOUND');
      updateRow_(entity, rowNum, rec);
    }
    return true;
  });
}

function getById_(entity, id) {
  const r = readTable(entity).filter(function (x) { return String(x[ID_FIELD[entity]]) === String(id); })[0];
  if (!r) throw new Error('NOT_FOUND');
  return r;
}

function assertCustomerVisible_(user, account) {
  if (!visibleCustomers_(user).some(function (c) { return c.account === account; })) throw new Error('NO_PERMISSION');
}

function saveUser_(rec, isNew) {
  if (!rec.username) throw new Error('REQUIRED_FIELDS');
  const users = readTable('Users');
  const clash = users.some(function (u) {
    return String(u.username).toLowerCase() === String(rec.username).toLowerCase() && u.userId !== rec.userId;
  });
  if (clash) throw new Error('DUPLICATE_ID');
  if (Array.isArray(rec.pages)) rec.pages = rec.pages.join(',');
  const pw = rec.password; delete rec.password;
  if (pw) { rec.salt = Utilities.getUuid(); rec.passwordHash = hash_(pw, rec.salt); }
  delete rec.lastLogin;
  if (isNew) {
    if (!pw) throw new Error('PASSWORD_REQUIRED');
    rec.userId = Utilities.getUuid(); rec.createdAt = new Date();
    appendRows('Users', [rec]);
  } else {
    const rowNum = findRow_('Users', rec.userId);
    if (rowNum < 0) throw new Error('NOT_FOUND');
    updateRow_('Users', rowNum, rec);
  }
  return true;
}

function deleteRecord(token, entity, id) {
  const page = ENTITY_PAGE[entity];
  const user = auth_(token, page);
  requireAction_(user, 'delete');
  return withLock_(function () {
    if (entity === 'Transactions') {
      const t = getById_(entity, id);
      assertCustomerVisible_(user, t.account);
      if (t.postStatus === 'Posted' && !user.isAdmin) throw new Error('POSTED_LOCKED');
      updateRow_(entity, t._row, { isDeleted: true, updatedByName: user.fullName || user.username, updatedAt: new Date() });
      return true;
    }
    if (entity === 'Customers') {
      assertCustomerVisible_(user, id);
      const used = readTable('Transactions').some(function (t) { return t.account === id && !t.isDeleted; });
      if (used) throw new Error('HAS_TRANSACTIONS');
    }
    if (entity === 'CustomerGroups' && readTable('Customers').some(function (c) { return c.group === id; })) {
      throw new Error('IN_USE');
    }
    if (entity === 'Users') {
      if (!user.isAdmin) throw new Error('NO_PERMISSION');
      if (id === user.userId) throw new Error('CANNOT_DELETE_SELF');
    }
    const rowNum = findRow_(entity, id);
    if (rowNum < 0) throw new Error('NOT_FOUND');
    sheet_(entity).deleteRow(rowNum);
    return true;
  });
}

/** Post (ترحيل) or un-post draft transactions. Only posted rows affect balances & reports. */
function postTransactions(token, ids, unpost) {
  const user = auth_(token, 'transactions');
  requireAction_(user, 'post');
  if (unpost && !user.isAdmin) throw new Error('NO_PERMISSION');
  return withLock_(function () {
    const set = {};
    ids.forEach(function (id) { set[id] = true; });
    const sh = sheet_('Transactions');
    const hm = headerMap_(sh);
    const idx = fieldIndex_('Transactions', hm);
    const last = sh.getLastRow();
    if (last < 2) return 0;
    const range = sh.getRange(2, 1, last - 1, hm.headers.length);
    const values = range.getValues();
    const allowed = {};
    visibleCustomers_(user).forEach(function (c) { allowed[c.account] = true; });
    let n = 0;
    const now = new Date(), who = user.fullName || user.username;
    values.forEach(function (row) {
      if (!set[String(row[idx.rowId])] || !allowed[String(row[idx.account]).trim()]) return;
      row[idx.postStatus] = unpost ? 'Draft' : 'Posted';
      row[idx.postedBy] = unpost ? '' : who;
      row[idx.postedAt] = unpost ? '' : now;
      n++;
    });
    range.setValues(values);
    return n;
  });
}

// ───────────────────────────── Import ─────────────────────────────
/**
 * Import rows parsed client-side (Excel/CSV from Dynamics 365).
 * rows: array of objects keyed by the file's headers. mode: 'append' | 'replace' | 'upsert'.
 */
function importRows(token, entity, rows, mode) {
  const user = auth_(token, 'import');
  requireAction_(user, 'add');
  if (mode === 'replace') requireAction_(user, 'delete');
  if (['Customers', 'Transactions', 'CustomerGroups'].indexOf(entity) === -1) throw new Error('Unknown entity');
  const norm = function (s) { return String(s).toLowerCase().replace(/[^a-z0-9؀-ۿ]/g, ''); };
  const aliases = {
    account: ['customeraccount', 'account', 'custaccount', 'accountnum', 'رقمالعميل'],
    name: ['name', 'customername', 'اسمالعميل'],
    group: ['customergroup', 'group', 'custgroup', 'groupid'],
    phone: ['contactnumber', 'phone', 'telephone', 'primarycontactphone'],
    email: ['email', 'primarycontactemail', 'emailaddress'],
    type: ['transactiontype', 'transtype', 'type'],
    amountCur: ['amountintransactioncurrency', 'amountcur', 'amount'],
    exchange: ['excange', 'exchange', 'exchangerate', 'exchrate'],
    creditMST: ['creditinlocalcurrency', 'creditmst', 'credit'],
    debitMST: ['debitinlocalcurrency', 'debitmst', 'debit'],
    createdAt: ['createddateandtime', 'createdat', 'createddatetime'],
    dueDate: ['duedate', 'due'],
    groupName: ['groupname', 'description']
  };
  const fieldFor = {};
  SCHEMA[entity].forEach(function (f) {
    fieldFor[norm(f[1])] = f[0];
    fieldFor[norm(f[0])] = f[0];
    (aliases[f[0]] || []).forEach(function (a) { if (!fieldFor[a]) fieldFor[a] = f[0]; });
  });
  const now = new Date(), who = user.fullName || user.username;
  const mapped = rows.map(function (src) {
    const o = {};
    Object.keys(src).forEach(function (h) {
      const n = norm(h);
      let k = fieldFor[n];
      if (!k) { // e.g. "Amount in transaction currency (- if Debit /+ if Credit)"
        const hit = Object.keys(fieldFor).filter(function (a) { return a.length >= 6 && n.indexOf(a) === 0; })
          .sort(function (a, b) { return b.length - a.length; })[0];
        k = hit && fieldFor[hit];
      }
      if (k && (o[k] === undefined || o[k] === '')) o[k] = src[h];
    });
    return o;
  }).filter(function (o) { return o[ID_FIELD[entity]] || (entity === 'Transactions' && o.account); });

  mapped.forEach(function (o) {
    if (entity === 'Transactions') {
      o.rowId = o.rowId || Utilities.getUuid();
      o.isDeleted = toBool_(o.isDeleted);
      o.postStatus = o.postStatus || 'Posted'; // data coming from D365 is already posted
      o.createdByName = o.createdByName || who;
      o.createdAt = o.createdAt || now;
      // Derive debit / credit from the signed amount when local columns are empty.
      if (!toNum_(o.debitMST) && !toNum_(o.creditMST) && o.amountCur !== undefined) {
        const amt = toNum_(o.amountCur) * (toNum_(o.exchange) || 1);
        if (amt < 0) o.debitMST = -amt; else o.creditMST = amt;
      }
      o.debitMST = Math.abs(toNum_(o.debitMST));
      o.creditMST = Math.abs(toNum_(o.creditMST));
    }
    if (entity === 'Customers') {
      o.status = o.status || 'Active';
      o.updatedBy = who; o.updatedAt = now;
    }
  });

  return withLock_(function () {
    const sh = sheet_(entity);
    if (mode === 'replace' && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    let inserted = 0, updated = 0;
    if (mode === 'upsert' && entity !== 'Transactions' && sh.getLastRow() > 1) {
      // One read + one write for the whole batch (fast for thousands of customers).
      const hm = headerMap_(sh), idx = fieldIndex_(entity, hm);
      const range = sh.getRange(2, 1, sh.getLastRow() - 1, hm.headers.length);
      const values = range.getValues();
      const pos = {};
      values.forEach(function (row, i) { pos[String(row[idx[ID_FIELD[entity]]]).trim()] = i; });
      const fresh = [];
      mapped.forEach(function (o) {
        const i = pos[String(o[ID_FIELD[entity]]).trim()];
        if (i === undefined) { fresh.push(o); return; }
        for (const k in o) if (idx[k] !== undefined && o[k] !== '' && o[k] !== undefined) values[i][idx[k]] = toSheetValue_(k, o[k]);
        updated++;
      });
      range.setValues(values);
      appendRows(entity, fresh); inserted = fresh.length;
    } else {
      appendRows(entity, mapped); inserted = mapped.length;
    }
    return { inserted: inserted, updated: updated, skipped: rows.length - mapped.length };
  });
}

/** Fills row_id for transaction rows pasted straight into the sheet so they can be edited. */
function ensureRowIds_() {
  const sh = sheet_('Transactions');
  const last = sh.getLastRow();
  if (last < 2) return;
  const col = fieldIndex_('Transactions', headerMap_(sh)).rowId + 1;
  const range = sh.getRange(2, col, last - 1, 1);
  const ids = range.getValues();
  let changed = false;
  ids.forEach(function (r) { if (!r[0]) { r[0] = Utilities.getUuid(); changed = true; } });
  if (changed) withLock_(function () { range.setValues(ids); });
}

function getTemplateHeaders(token, entity) {
  auth_(token, 'import');
  const system = ['rowId', 'createdById', 'createdAt', 'updatedByName', 'updatedAt', 'postStatus', 'postedBy', 'postedAt', 'updatedBy'];
  return SCHEMA[entity].filter(function (f) { return system.indexOf(f[0]) === -1; }).map(function (f) { return f[1]; });
}

function saveSettings(token, values) {
  const user = auth_(token, 'users');
  if (!user.isAdmin) throw new Error('NO_PERMISSION');
  return withLock_(function () {
    const rows = readTable('Settings');
    Object.keys(values).forEach(function (k) {
      const r = rows.filter(function (x) { return x.key === k; })[0];
      if (r) updateRow_('Settings', r._row, { value: values[k] });
      else appendRows('Settings', [{ key: k, value: values[k] }]);
    });
    return getSettings_();
  });
}

// ───────────────────────────── Ledger engine ─────────────────────────────
/**
 * Loads posted, non-deleted transactions for visible customers and settles credits
 * against debits FIFO (oldest invoice first) to get open amounts, paid dates and aging.
 */
function ledger_(user, asOf) {
  const custs = visibleCustomers_(user);
  const cmap = {};
  custs.forEach(function (c) { cmap[c.account] = c; });
  const tx = readTable('Transactions').filter(function (t) {
    return !t.isDeleted && t.postStatus !== 'Draft' && cmap[t.account] && t.date && (!asOf || t.date <= asOf);
  });
  const byCust = {};
  tx.forEach(function (t) { (byCust[t.account] = byCust[t.account] || []).push(t); });
  const invoices = [];
  Object.keys(byCust).forEach(function (acc) {
    const list = byCust[acc].sort(function (a, b) {
      return a.date < b.date ? -1 : a.date > b.date ? 1 : (b.debitMST - a.debitMST);
    });
    const queue = [];
    let unapplied = 0;
    const terms = +cmap[acc].termsDays || 0;
    list.forEach(function (t) {
      const net = round2_((+t.debitMST || 0) - (+t.creditMST || 0));
      if (net > 0) {
        const inv = {
          account: acc, voucher: t.voucher, type: t.type, date: t.date, description: t.description,
          dueDate: t.dueDate || addDays_(t.date, terms), amount: net, remaining: net, paidDate: '', lastPaymentDate: ''
        };
        if (unapplied > 0) {
          const a = Math.min(unapplied, inv.remaining);
          inv.remaining = round2_(inv.remaining - a); unapplied = round2_(unapplied - a);
          inv.lastPaymentDate = t.date;
          if (inv.remaining <= 0.005) inv.paidDate = t.date;
        }
        invoices.push(inv);
        if (inv.remaining > 0.005) queue.push(inv);
      } else if (net < 0) {
        let credit = -net;
        while (credit > 0.005 && queue.length) {
          const inv = queue[0];
          const a = Math.min(credit, inv.remaining);
          inv.remaining = round2_(inv.remaining - a); credit = round2_(credit - a);
          inv.lastPaymentDate = t.date;
          if (inv.remaining <= 0.005) { inv.remaining = 0; inv.paidDate = t.date; queue.shift(); }
        }
        unapplied = round2_(unapplied + credit);
      }
    });
  });
  return { customers: custs, cmap: cmap, tx: tx, invoices: invoices };
}

function addDays_(ymd, days) {
  const d = toDate_(ymd);
  d.setDate(d.getDate() + days);
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

function daysBetween_(fromYmd, toYmd) {
  return Math.round((toDate_(toYmd) - toDate_(fromYmd)) / 86400000);
}

function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

function customerFilter_(c, f) {
  if (f.group && c.group !== f.group) return false;
  if (f.collector && c.collector !== f.collector) return false;
  if (f.account && c.account !== f.account) return false;
  if (f.status && c.status !== f.status) return false;
  return true;
}

function bucket_(days) {
  return days <= 30 ? 'b1' : days <= 60 ? 'b2' : days <= 90 ? 'b3' : 'b4';
}

// ───────────────────────────── Reports ─────────────────────────────
/** Opening balance, period debit/credit and closing balance per customer. */
function reportBalances(token, f) {
  const user = auth_(token, 'balances');
  f = f || {};
  const from = f.dateFrom || '1900-01-01', to = f.dateTo || today_();
  const L = ledger_(user, to);
  const agg = {};
  L.tx.forEach(function (t) {
    const a = agg[t.account] = agg[t.account] || { opening: 0, debit: 0, credit: 0 };
    const net = (+t.debitMST || 0) - (+t.creditMST || 0);
    if (t.date < from) a.opening += net;
    else { a.debit += +t.debitMST || 0; a.credit += +t.creditMST || 0; }
  });
  const rows = L.customers.filter(function (c) { return customerFilter_(c, f); }).map(function (c) {
    const a = agg[c.account] || { opening: 0, debit: 0, credit: 0 };
    return {
      account: c.account, name: c.name, group: c.group, collector: c.collector, currency: c.currency,
      opening: round2_(a.opening), debit: round2_(a.debit), credit: round2_(a.credit),
      closing: round2_(a.opening + a.debit - a.credit)
    };
  }).filter(function (r) { return !f.hideZero || r.opening || r.debit || r.credit || r.closing; })
    .sort(function (a, b) { return b.closing - a.closing; });
  return {
    rows: rows, from: from, to: to,
    totals: { opening: sum_(rows, 'opening'), debit: sum_(rows, 'debit'), credit: sum_(rows, 'credit'), closing: sum_(rows, 'closing') }
  };
}

/** Customer statement of account with running balance. */
function customerStatement(token, account, dateFrom, dateTo) {
  const user = auth_(token);
  const pages = ['customers', 'balances', 'paid', 'overdue', 'followup', 'collection', 'dashboard'];
  if (!pages.some(function (p) { return user.pageList.indexOf(p) > -1; })) throw new Error('NO_PERMISSION');
  assertCustomerVisible_(user, account);
  const from = dateFrom || '1900-01-01', to = dateTo || today_();
  const L = ledger_(user, to);
  const cust = L.cmap[account];
  let opening = 0;
  const lines = [];
  L.tx.filter(function (t) { return t.account === account; })
    .sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; })
    .forEach(function (t) {
      const net = (+t.debitMST || 0) - (+t.creditMST || 0);
      if (t.date < from) { opening += net; return; }
      lines.push({ date: t.date, voucher: t.voucher, type: t.type, description: t.description, dueDate: t.dueDate, debit: t.debitMST, credit: t.creditMST });
    });
  let run = opening;
  lines.forEach(function (l) { run = round2_(run + (+l.debit || 0) - (+l.credit || 0)); l.balance = run; });
  const today = today_();
  const open = L.invoices.filter(function (i) { return i.account === account && i.remaining > 0; });
  const overdue = open.filter(function (i) { return i.dueDate < today; });
  return {
    customer: cust, opening: round2_(opening), closing: round2_(run), lines: lines,
    openAmount: sum_(open, 'remaining'), overdueAmount: sum_(overdue, 'remaining'),
    maxDays: overdue.reduce(function (m, i) { return Math.max(m, daysBetween_(i.dueDate, today)); }, 0),
    followUps: readTable('FollowUps').filter(function (x) { return x.account === account; })
      .sort(function (a, b) { return a.actionDate < b.actionDate ? 1 : -1; }).slice(0, 20)
  };
}

/** Invoices fully settled within the period. */
function reportPaidInvoices(token, f) {
  const user = auth_(token, 'paid');
  f = f || {};
  const from = f.dateFrom || '1900-01-01', to = f.dateTo || today_();
  const L = ledger_(user, to);
  const rows = L.invoices.filter(function (i) {
    return i.paidDate && i.paidDate >= from && i.paidDate <= to && customerFilter_(L.cmap[i.account], f);
  }).map(function (i) {
    const c = L.cmap[i.account];
    return {
      account: i.account, name: c.name, collector: c.collector, voucher: i.voucher, date: i.date,
      dueDate: i.dueDate, paidDate: i.paidDate, amount: i.amount,
      lateDays: Math.max(0, daysBetween_(i.dueDate, i.paidDate))
    };
  }).sort(function (a, b) { return a.paidDate < b.paidDate ? 1 : -1; });
  const onTime = rows.filter(function (r) { return r.lateDays === 0; }).length;
  return { rows: rows, totals: { amount: sum_(rows, 'amount'), count: rows.length, onTime: onTime, late: rows.length - onTime } };
}

/** Open invoices past their expected collection (due) date. */
function reportOverdue(token, f) {
  const user = auth_(token, 'overdue');
  return overdue_(user, f || {});
}

function overdue_(user, f) {
  const asOf = f.asOf || today_();
  const L = ledger_(user, asOf);
  const minDays = +f.minDays || 1;
  const lastFollow = {};
  readTable('FollowUps').forEach(function (x) {
    if (!lastFollow[x.account] || x.actionDate > lastFollow[x.account].actionDate) lastFollow[x.account] = x;
  });
  const rows = L.invoices.filter(function (i) {
    return i.remaining > 0 && i.dueDate < asOf && customerFilter_(L.cmap[i.account], f);
  }).map(function (i) {
    const c = L.cmap[i.account];
    const days = daysBetween_(i.dueDate, asOf);
    return {
      account: i.account, name: c.name, group: c.group, collector: c.collector, phone: c.phone, email: c.email,
      voucher: i.voucher, date: i.date, dueDate: i.dueDate, amount: i.amount, remaining: i.remaining,
      days: days, bucket: bucket_(days), lastPaymentDate: i.lastPaymentDate,
      lastFollowUp: lastFollow[i.account] ? lastFollow[i.account].actionDate : '',
      promiseDate: lastFollow[i.account] ? lastFollow[i.account].promiseDate : ''
    };
  }).filter(function (r) { return r.days >= minDays; })
    .sort(function (a, b) { return b.days - a.days; });
  const buckets = { b1: 0, b2: 0, b3: 0, b4: 0 };
  rows.forEach(function (r) { buckets[r.bucket] = round2_(buckets[r.bucket] + r.remaining); });
  // Per customer summary
  const byCust = {};
  rows.forEach(function (r) {
    const s = byCust[r.account] = byCust[r.account] || {
      account: r.account, name: r.name, collector: r.collector, phone: r.phone, email: r.email,
      count: 0, remaining: 0, maxDays: 0, lastFollowUp: r.lastFollowUp, promiseDate: r.promiseDate
    };
    s.count++; s.remaining = round2_(s.remaining + r.remaining); s.maxDays = Math.max(s.maxDays, r.days);
  });
  const customers = Object.keys(byCust).map(function (k) { return byCust[k]; })
    .sort(function (a, b) { return b.remaining - a.remaining; });
  return {
    asOf: asOf, rows: rows, customers: customers, buckets: buckets,
    totals: { remaining: sum_(rows, 'remaining'), count: rows.length, customers: customers.length,
      avgDays: rows.length ? Math.round(rows.reduce(function (s, r) { return s + r.days; }, 0) / rows.length) : 0 }
  };
}

/** Follow-up page: overdue customers + their follow-up log. */
function getFollowUpBoard(token, f) {
  const user = auth_(token, 'followup');
  const od = overdue_(user, f || {});
  const today = today_();
  const fu = readTable('FollowUps');
  const due = fu.filter(function (x) { return x.status === 'Open' && x.promiseDate && x.promiseDate <= today; });
  return { customers: od.customers, totals: od.totals, promisesDue: due.length };
}

/** Collection team dashboard: per collector → companies, overdue, days late. */
function reportCollection(token, f) {
  const user = auth_(token, 'collection');
  f = f || {};
  const od = overdue_(user, { asOf: f.asOf, group: f.group });
  const L = ledger_(user, f.asOf || today_());
  const users = {};
  readTable('Users').forEach(function (u) { users[u.username] = u.fullName || u.username; });
  const balance = {};
  L.tx.forEach(function (t) { balance[t.account] = (balance[t.account] || 0) + (+t.debitMST || 0) - (+t.creditMST || 0); });
  const monthStart = (f.asOf || today_()).slice(0, 8) + '01';
  const collected = {};
  L.tx.forEach(function (t) {
    if (t.date >= monthStart && +t.creditMST > 0) collected[t.account] = (collected[t.account] || 0) + (+t.creditMST);
  });
  const team = {};
  L.customers.filter(function (c) { return !f.group || c.group === f.group; }).forEach(function (c) {
    const key = c.collector || '';
    const m = team[key] = team[key] || {
      collector: key, name: key ? (users[key] || key) : '', customers: 0, balance: 0, overdue: 0,
      overdueInvoices: 0, overdueCustomers: 0, maxDays: 0, sumDays: 0, collectedMTD: 0, companies: []
    };
    m.customers++;
    m.balance = round2_(m.balance + (balance[c.account] || 0));
    m.collectedMTD = round2_(m.collectedMTD + (collected[c.account] || 0));
  });
  od.customers.forEach(function (oc) {
    const m = team[oc.collector || ''];
    if (!m) return;
    m.overdue = round2_(m.overdue + oc.remaining);
    m.overdueInvoices += oc.count;
    m.overdueCustomers++;
    m.maxDays = Math.max(m.maxDays, oc.maxDays);
    m.companies.push({ account: oc.account, name: oc.name, remaining: oc.remaining, maxDays: oc.maxDays, count: oc.count, lastFollowUp: oc.lastFollowUp, balance: round2_(balance[oc.account] || 0) });
  });
  od.rows.forEach(function (r) { const m = team[r.collector || '']; if (m) m.sumDays += r.days; });
  const rows = Object.keys(team).map(function (k) {
    const m = team[k];
    m.avgDays = m.overdueInvoices ? Math.round(m.sumDays / m.overdueInvoices) : 0;
    m.companies.sort(function (a, b) { return b.maxDays - a.maxDays; });
    delete m.sumDays;
    return m;
  }).filter(function (m) { return !f.collector || m.collector === f.collector; })
    .sort(function (a, b) { return b.overdue - a.overdue; });
  return { rows: rows, buckets: od.buckets, totals: { overdue: sum_(rows, 'overdue'), balance: sum_(rows, 'balance'), collectedMTD: sum_(rows, 'collectedMTD') } };
}

/** Main dashboard KPIs. */
function getDashboard(token) {
  const user = auth_(token, 'dashboard');
  const today = today_();
  const L = ledger_(user, today);
  const od = overdue_(user, { asOf: today });
  const receivable = L.tx.reduce(function (s, t) { return s + (+t.debitMST || 0) - (+t.creditMST || 0); }, 0);
  const monthStart = today.slice(0, 8) + '01';
  const months = [];
  const base = toDate_(monthStart);
  for (let i = 5; i >= 0; i--) {
    const d = new Date(base.getFullYear(), base.getMonth() - i, 1);
    months.push({ key: Utilities.formatDate(d, TZ, 'yyyy-MM'), label: Utilities.formatDate(d, TZ, 'MMM'), debit: 0, credit: 0 });
  }
  const mIdx = {};
  months.forEach(function (m, i) { mIdx[m.key] = i; });
  let collectedMTD = 0;
  L.tx.forEach(function (t) {
    const i = mIdx[t.date.slice(0, 7)];
    if (i !== undefined) { months[i].debit += +t.debitMST || 0; months[i].credit += +t.creditMST || 0; }
    if (t.date >= monthStart) collectedMTD += +t.creditMST || 0;
  });
  months.forEach(function (m) { m.debit = round2_(m.debit); m.credit = round2_(m.credit); });
  const drafts = readTable('Transactions').filter(function (t) { return !t.isDeleted && t.postStatus === 'Draft' && L.cmap[t.account]; }).length;
  const balances = {};
  L.tx.forEach(function (t) { balances[t.account] = (balances[t.account] || 0) + (+t.debitMST || 0) - (+t.creditMST || 0); });
  const top = Object.keys(balances).map(function (a) {
    const c = L.cmap[a];
    const o = od.customers.filter(function (x) { return x.account === a; })[0];
    return { account: a, name: c.name, collector: c.collector, group: c.group, balance: round2_(balances[a]),
      overdue: o ? o.remaining : 0, maxDays: o ? o.maxDays : 0 };
  }).sort(function (a, b) { return b.balance - a.balance; }).slice(0, 10);
  return {
    kpi: {
      customers: L.customers.length,
      active: L.customers.filter(function (c) { return c.status !== 'Blocked' && c.status !== 'Inactive'; }).length,
      receivable: round2_(receivable), overdue: od.totals.remaining, overdueCustomers: od.totals.customers,
      collectedMTD: round2_(collectedMTD), drafts: drafts, avgDays: od.totals.avgDays
    },
    months: months, buckets: od.buckets, top: top
  };
}

// ───────────────────────────── Demo data ─────────────────────────────
function loadDemoData() {
  setup();
  const groups = [{ groupId: 'CORP', groupName: 'Corporate' }, { groupId: 'GOV', groupName: 'Government' }, { groupId: 'RET', groupName: 'Retail' }];
  appendRows('CustomerGroups', groups);
  const salt = Utilities.getUuid();
  appendRows('Users', [{
    userId: Utilities.getUuid(), username: 'collector1', fullName: 'Ahmed Collector', email: '',
    passwordHash: hash_('collector123', salt), salt: salt, role: 'User',
    pages: 'dashboard,customers,overdue,followup,collection', canAdd: true, canEdit: true,
    canDelete: false, canPost: false, isCollector: true, scope: 'own', active: true, createdAt: new Date()
  }]);
  const names = ['Al Noor Trading', 'Gulf Steel Co.', 'Riyadh Medical', 'Desert Foods', 'Blue Sea Logistics', 'Najd Contracting', 'Ministry Supplies', 'Smart Retail'];
  const custs = names.map(function (n, i) {
    return {
      account: 'C' + (1001 + i), name: n, group: groups[i % 3].groupId, currency: 'SAR', contactPerson: 'Contact ' + (i + 1),
      phone: '05' + (50000000 + i * 1111), email: 'ar' + (i + 1) + '@example.com', address: 'Street ' + (i + 10),
      city: ['Riyadh', 'Jeddah', 'Dammam'][i % 3], country: 'SA', creditLimit: 100000, termsDays: 30,
      collector: i % 2 ? 'collector1' : 'admin', status: 'Active', createdAt: new Date(2025, 0, 1 + i)
    };
  });
  appendRows('Customers', custs);
  const tx = [];
  const now = new Date();
  custs.forEach(function (c, ci) {
    for (let k = 0; k < 8; k++) {
      const d = new Date(now.getFullYear(), now.getMonth() - 7 + k, 3 + ci);
      const amt = 5000 + ((ci * 37 + k * 53) % 20) * 1000;
      tx.push({ voucher: 'INV-' + c.account + '-' + k, type: 'Invoice', account: c.account, date: d, description: 'Sales invoice',
        dueDate: new Date(d.getTime() + 30 * 86400000), amountCur: -amt, currency: 'SAR', exchange: 1, debitMST: amt, creditMST: 0 });
      if (k < 8 - (ci % 4) - 1) {
        const p = new Date(d.getTime() + (20 + ci * 6) * 86400000);
        if (p < now) tx.push({ voucher: 'PAY-' + c.account + '-' + k, type: 'Payment', account: c.account, date: p, description: 'Customer payment',
          amountCur: amt, currency: 'SAR', exchange: 1, debitMST: 0, creditMST: amt });
      }
    }
  });
  tx.forEach(function (t) {
    t.rowId = Utilities.getUuid(); t.postStatus = 'Posted'; t.isDeleted = false;
    t.createdByName = 'D365 import'; t.createdAt = new Date(); t.userAdd = 'admin';
  });
  appendRows('Transactions', tx);
  return 'Demo data loaded';
}
