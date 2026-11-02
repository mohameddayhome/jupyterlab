/*******************************************************
 * نظام إعداد القوائم المالية - مكتب المدقق
 * (نسخة محسّنة: كتابة مجمّعة أسرع + قفل للتزامن + إصلاحات)
 *******************************************************/

const SHEET_NAMES = {
  CLIENTS : 'العملاء',
  USERS   : 'المستخدمين',
  COA     : 'دليل_الحسابات',
  MAPPING : 'ربط_الحسابات',
  TB      : 'ميزان_المراجعة',
  BS      : 'الميزانية_العمومية',
  IS      : 'قائمة_الدخل',
  CF      : 'قائمة_التدفقات',
  RATIOS  : 'مؤشرات_الأداء',
  WP      : 'ورقة_العمل'
};

const WP_HEADERS = ['ClientID','السنة المالية','القائمة','بند القوائم','رقم الحساب','اسم الحساب','الحساب المعياري','المبلغ'];


// ألوان رؤوس الجداول بنمط Dynamics 365
const HEADER_BG = '#002050';
const HEADER_FG = '#FFFFFF';

const COA_HEADERS = ['كود الحساب','اسم الحساب','نوع الحساب','بند القائمة','تصنيف','الرصيد الطبيعي','تصنيف التدفقات','ترتيب'];

const COA_SEED = [
  ['1101','النقدية وما يعادلها', 'أصول','الميزانية','متداول','مدين','تشغيلي',110],
  ['1102','المدينون (العملاء)', 'أصول','الميزانية','متداول','مدين','تشغيلي',120],
  ['1103','المخزون', 'أصول','الميزانية','متداول','مدين','تشغيلي',130],
  ['1104','مصروفات مدفوعة مقدماً', 'أصول','الميزانية','متداول','مدين','تشغيلي',140],
  ['1201','الأصول الثابتة (تكلفة)', 'أصول','الميزانية','غير متداول','مدين','استثماري',210],
  ['1202','مجمع إهلاك الأصول الثابتة', 'أصول','الميزانية','غير متداول','دائن','',220],
  ['1203','الأصول غير الملموسة', 'أصول','الميزانية','غير متداول','مدين','استثماري',230],
  ['2101','الدائنون (الموردون)', 'التزامات','الميزانية','متداول','دائن','تشغيلي',310],
  ['2102','قروض قصيرة الأجل', 'التزامات','الميزانية','متداول','دائن','تمويلي',320],
  ['2103','مصروفات مستحقة', 'التزامات','الميزانية','متداول','دائن','تشغيلي',330],
  ['2201','قروض طويلة الأجل', 'التزامات','الميزانية','غير متداول','دائن','تمويلي',410],
  ['3001','رأس المال', 'حقوق ملكية','الميزانية','دائم','دائن','تمويلي',510],
  ['3002','الاحتياطيات', 'حقوق ملكية','الميزانية','دائم','دائن','',520],
  ['3003','الأرباح المحتجزة', 'حقوق ملكية','الميزانية','دائم','دائن','',530],
  ['4001','إيرادات النشاط', 'إيرادات','قائمة الدخل','إيراد تشغيلي','دائن','',610],
  ['4002','إيرادات أخرى', 'إيرادات','قائمة الدخل','إيراد آخر','دائن','',620],
  ['5001','تكلفة المبيعات', 'مصروفات','قائمة الدخل','تكلفة','مدين','',710],
  ['6001','مصروفات إدارية وعمومية', 'مصروفات','قائمة الدخل','مصروف تشغيلي','مدين','',720],
  ['6002','مصروفات بيعية وتسويقية', 'مصروفات','قائمة الدخل','مصروف تشغيلي','مدين','',730],
  ['6003','مصروفات مالية (فوائد)', 'مصروفات','قائمة الدخل','مصروف مالي','مدين','',740],
  ['6004','مصروف الإهلاك', 'مصروفات','قائمة الدخل','مصروف تشغيلي','مدين','',750]
];

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('نظام إعداد القوائم المالية')
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 القوائم المالية')
    .addItem('⚙️ إعداد جداول النظام', 'setupWorkbook')
    .addSeparator()
    .addItem('🌐 فتح واجهة التطبيق', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('نظام القوائم المالية').setWidth(1400).setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'نظام القوائم المالية');
}

/* دالة آمنة للبحث عن الشيت */
function getSheetByNameSafe_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim() === name.trim()) return sheets[i];
  }
  return null;
}

/* يرجع الشيت أو يرمي خطأ واضح للمستخدم */
function requireSheet_(name) {
  const sh = getSheetByNameSafe_(SpreadsheetApp.getActiveSpreadsheet(), name);
  if (!sh) throw new Error('الجدول "' + name + '" غير موجود. اضغط "إعداد قاعدة البيانات" أولاً.');
  return sh;
}

function setupWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return "خطأ: لم يتم العثور على ملف جوجل شيت.";

  ensureSheet_(ss, SHEET_NAMES.CLIENTS, ['ClientID','اسم العميل','الرقم الضريبي','بداية السنة المالية','نهاية السنة المالية','العملة','حالة العميل','ملاحظات']);

  const usersSh = ensureSheet_(ss, SHEET_NAMES.USERS, ['اسم المستخدم','كلمة المرور','الاسم','الصلاحية','الحالة']);
  if (usersSh.getLastRow() < 2) {
    usersSh.appendRow(['admin', 'admin123', 'المدير العام', 'مدير', 'نشط']);
  }

  const coaSh = ensureSheet_(ss, SHEET_NAMES.COA, COA_HEADERS);
  if (coaSh.getLastRow() < 2) {
    coaSh.getRange(2, 1, COA_SEED.length, COA_SEED[0].length).setValues(COA_SEED);
  }

  ensureSheet_(ss, SHEET_NAMES.MAPPING, ['ClientID','كود الحساب بالشركة','اسم الحساب بالشركة','كود الحساب المعياري','اسم الحساب المعياري','ملاحظات']);
  ensureSheet_(ss, SHEET_NAMES.TB, ['ClientID','السنة المالية','كود الحساب','اسم الحساب','مدين','دائن','ملاحظات']);
  ensureSheet_(ss, SHEET_NAMES.BS, ['ClientID','السنة المالية','القسم','البند','المبلغ']);
  ensureSheet_(ss, SHEET_NAMES.IS, ['ClientID','السنة المالية','البند','المبلغ']);
  ensureSheet_(ss, SHEET_NAMES.CF, ['ClientID','السنة المالية','القسم','البند','المبلغ','النوع']);
  ensureSheet_(ss, SHEET_NAMES.RATIOS, ['ClientID','السنة المالية','المؤشر','القيمة','التقييم']);
  ensureSheet_(ss, SHEET_NAMES.WP, WP_HEADERS);

  return '✅ تم إعداد جميع الجداول وإنشاء المستخدم الافتراضي (admin / admin123)';
}

function ensureSheet_(ss, name, headers) {
  let sh = getSheetByNameSafe_(ss, name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground(HEADER_BG).setFontColor(HEADER_FG)
      .setFontFamily('Segoe UI');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
    sh.autoResizeColumns(1, headers.length);
  }
  return sh;
}

/* ================== أداة كتابة مجمّعة (أسرع بكثير من deleteRow/appendRow) ================== */
/**
 * يحذف الصفوف التي تحقق الشرط ثم يضيف الصفوف الجديدة، بعملية قراءة وكتابة واحدة.
 * @param {Sheet} sh
 * @param {function(Array):boolean} shouldRemove
 * @param {Array<Array>} newRows
 */
function replaceRows_(sh, shouldRemove, newRows) {
  if (!sh) throw new Error('الجدول غير موجود. يرجى إعداد قاعدة البيانات.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const lastRow = sh.getLastRow();
    const lastCol = Math.max(sh.getLastColumn(), (newRows && newRows[0]) ? newRows[0].length : 1);
    let kept = [], footer = [];
    if (lastRow > 1) {
      const range = sh.getRange(2, 1, lastRow - 1, lastCol);
      const vals = range.getValues(), forms = range.getFormulas();
      vals.forEach((r, i) => {
        const row = r.map((v, j) => forms[i][j] || v);   // الحفاظ على معادلات المستخدم
        if (String(r[0]).trim() === '') {
          // صفوف بدون ClientID (مثل صف إجماليات أضافه المستخدم) تبقى في آخر الجدول
          if (r.some(x => x !== '')) footer.push(row);
        } else if (!shouldRemove(r)) {
          kept.push(row);
        }
      });
    }
    const added = (newRows || []).map(r => {
      const a = r.slice(0, lastCol);
      while (a.length < lastCol) a.push('');
      return a;
    });
    const all = kept.concat(added, footer);
    if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    if (all.length) sh.getRange(2, 1, all.length, lastCol).setValues(all);
  } finally {
    lock.releaseLock();
  }
}

/* ================== المصادقة ================== */
function login(username, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, msg: 'لم يتم العثور على الشيت.' };

    const sh = getSheetByNameSafe_(ss, SHEET_NAMES.USERS);
    if (!sh) return { success: false, msg: 'جدول المستخدمين غير موجود. اضغط إعداد قاعدة البيانات أولاً.' };

    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const dbUser = String(data[i][0]).trim();
      const dbPass = String(data[i][1]).trim();
      const dbStatus = String(data[i][4]).trim();

      if (dbUser === String(username).trim() && dbPass === String(password).trim()) {
        if (dbStatus === 'نشط' || dbStatus === '') {
          const user = { username: dbUser, name: String(data[i][2]), role: String(data[i][3]) };
          const token = Utilities.getUuid();
          CacheService.getScriptCache().put(SESSION_PREFIX + token, JSON.stringify(user), SESSION_IDLE_SECONDS);
          return { success: true, token: token, username: user.username, name: user.name, role: user.role };
        }
        return { success: false, msg: 'حسابك موقوف. يرجى التواصل مع المدير.' };
      }
    }
    return { success: false, msg: 'اسم المستخدم أو كلمة المرور غير صحيحة.' };
  } catch(e) {
    return { success: false, msg: 'خطأ داخلي: ' + e.message };
  }
}

/* ================== الجلسات ================== */
// الجلسة تنتهي بعد 30 دقيقة بدون نشاط، وتتجدد مع كل عملية
const SESSION_PREFIX = 'sess_';
const SESSION_IDLE_SECONDS = 30 * 60;

function getSession_(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const raw = cache.get(SESSION_PREFIX + token);
  if (!raw) return null;
  cache.put(SESSION_PREFIX + token, raw, SESSION_IDLE_SECONDS);
  return JSON.parse(raw);
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove(SESSION_PREFIX + token);
  return true;
}

/**
 * نقطة الدخول الوحيدة للواجهة: تتحقق من الجلسة ثم تنفّذ العملية المطلوبة.
 * دوال البيانات تنتهي بـ "_" فلا يمكن استدعاؤها مباشرة من المتصفح.
 */
function api(token, fn, args) {
  if (!getSession_(token)) throw new Error('SESSION_EXPIRED');
  const routes = {
    getCOA: getCOA_, saveCOAAccount: saveCOAAccount_, deleteCOAAccount: deleteCOAAccount_,
    getDashboardStats: getDashboardStats_,
    getClients: getClients_, saveClient: saveClient_, deleteClient: deleteClient_,
    getTrialBalance: getTrialBalance_, saveTrialBalance: saveTrialBalance_,
    getMapping: getMapping_, getMappingView: getMappingView_, saveMapping: saveMapping_, autoMapClient: autoMapClient_,
    computeStatements: computeStatements_, getSavedStatements: getSavedStatements_,
    getComparisonData: getComparisonData_,
    getWorkingPaper: getWorkingPaper_, checkTrialBalanceDuplicate: checkTrialBalanceDuplicate_
  };
  const f = routes[fn];
  if (!f) throw new Error('عملية غير مسموحة: ' + fn);
  return f.apply(null, args || []);
}

/* ================== دليل الحسابات (COA) ================== */
function getCOA_() {
  const sh = getSheetByNameSafe_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_NAMES.COA);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues().filter(r => r[0]).map(r => ({
      code: String(r[0]), name: String(r[1]), type: r[2], stmt: r[3], classif: r[4], normal: r[5], cfCat: r[6], sort: Number(r[7]) || 0
  }));
}

/* acc.originalCode (اختياري): الكود القديم عند تعديل كود الحساب نفسه */
function saveCOAAccount_(acc) {
  const sh = requireSheet_(SHEET_NAMES.COA);
  const data = sh.getDataRange().getValues();
  const lookup = String(acc.originalCode || acc.code);
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === lookup) { rowIdx = i + 1; break; }
  }
  // منع تكرار الكود عند الإضافة أو تغيير الكود
  if (String(acc.code) !== lookup || rowIdx < 0) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(acc.code) && (i + 1) !== rowIdx) {
        throw new Error('كود الحساب ' + acc.code + ' موجود مسبقاً.');
      }
    }
  }
  const row = [acc.code, acc.name, acc.type, acc.stmt, acc.classif, acc.normal, acc.cfCat, acc.sort];
  if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  return true;
}

function deleteCOAAccount_(code) {
  const sh = requireSheet_(SHEET_NAMES.COA);
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(code)) { sh.deleteRow(i + 1); return true; }
  }
  return false;
}

/* ================== أدوات مساعدة ================== */
function fmtDate_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}

function isClientYear_(r, clientId, year) {
  return String(r[0]) === String(clientId) && String(r[1]) === String(year);
}

function getDashboardStats_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const count = n => { const s = getSheetByNameSafe_(ss, n); return s ? Math.max(0, s.getLastRow() - 1) : 0; };

  // عدد القوائم = عدد (عميل + سنة) الفريدة وليس عدد الصفوف
  let statements = 0;
  const bs = getSheetByNameSafe_(ss, SHEET_NAMES.BS);
  if (bs && bs.getLastRow() > 1) {
    const keys = {};
    bs.getRange(2, 1, bs.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0]) keys[r[0] + '|' + r[1]] = 1; });
    statements = Object.keys(keys).length;
  }
  return { clients: count(SHEET_NAMES.CLIENTS), tbs: count(SHEET_NAMES.TB), statements: statements, coa: count(SHEET_NAMES.COA) };
}

/* ================== العملاء ================== */
function getClients_() {
  const sh = getSheetByNameSafe_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_NAMES.CLIENTS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues().filter(r => r[0]).map(r => ({
      id: String(r[0]), name: r[1], tax: r[2], startDate: fmtDate_(r[3]), endDate: fmtDate_(r[4]),
      currency: r[5], status: r[6], notes: r[7]
  }));
}

function saveClient_(c) {
  const sh = requireSheet_(SHEET_NAMES.CLIENTS);
  if (!c.id) c.id = 'CL' + Date.now();
  const data = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) if (String(data[i][0]) === String(c.id)) { rowIdx = i + 1; break; }
  const row = [c.id, c.name||'', c.tax||'', c.startDate||'', c.endDate||'', c.currency||'SAR', c.status||'نشط', c.notes||''];
  if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  return c.id;
}

function deleteClient_(id) {
  const sh = requireSheet_(SHEET_NAMES.CLIENTS);
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) if (String(data[i][0]) === String(id)) { sh.deleteRow(i + 1); break; }
  return true;
}

/* ================== ميزان المراجعة والربط ================== */
function getTrialBalance_(clientId, year) {
  const sh = getSheetByNameSafe_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_NAMES.TB);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues()
    .filter(r => isClientYear_(r, clientId, year))
    .map(r => ({ code: String(r[2]), name: String(r[3]), debit: Number(r[4]) || 0, credit: Number(r[5]) || 0, notes: r[6] || '' }));
}

function saveTrialBalance_(clientId, year, rows) {
  const sh = requireSheet_(SHEET_NAMES.TB);
  const newRows = (rows || []).map(r => [clientId, year, r.code, r.name, Number(r.debit)||0, Number(r.credit)||0, r.notes||'']);
  replaceRows_(sh, r => isClientYear_(r, clientId, year), newRows);
  return true;
}

/* بصمة الميزان: الأكواد والمبالغ مرتبة، لاكتشاف رفع نفس الملف مرتين */
function tbFingerprint_(rows) {
  return (rows || [])
    .filter(r => String(r.code || '').trim())
    .map(r => [String(r.code).trim(), Math.round((Number(r.debit) || 0) * 100), Math.round((Number(r.credit) || 0) * 100)].join('|'))
    .sort().join(';');
}

/**
 * فحص التكرار قبل حفظ ميزان المراجعة:
 * - existingRows: يوجد ميزان محفوظ لنفس العميل والسنة (سيُستبدل)
 * - identicalToExisting: الميزان الجديد مطابق تماماً للمحفوظ
 * - sameAs: مطابق لميزان سنة أخرى أو عميل آخر (غالباً نفس الملف مرفوع بالخطأ)
 * - dupCodes: أكواد مكررة داخل الميزان نفسه (ستُجمع أرصدتها)
 */
function checkTrialBalanceDuplicate_(clientId, year, rows) {
  const result = { existingRows: 0, identicalToExisting: false, sameAs: [], dupCodes: [] };

  const seen = {};
  (rows || []).forEach(r => { const c = String(r.code || '').trim(); if (c) seen[c] = (seen[c] || 0) + 1; });
  result.dupCodes = Object.keys(seen).filter(c => seen[c] > 1);

  const sh = getSheetByNameSafe_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_NAMES.TB);
  if (!sh || sh.getLastRow() < 2) return result;

  const groups = {};
  sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues().forEach(r => {
    if (!r[0]) return;
    const k = r[0] + '|' + r[1];
    if (!groups[k]) groups[k] = { clientId: String(r[0]), year: String(r[1]), rows: [] };
    groups[k].rows.push({ code: r[2], debit: r[4], credit: r[5] });
  });

  const fp = tbFingerprint_(rows);
  const names = {};
  getClients_().forEach(c => names[c.id] = c.name);

  Object.keys(groups).forEach(k => {
    const g = groups[k];
    const same = fp !== '' && tbFingerprint_(g.rows) === fp;
    if (g.clientId === String(clientId) && g.year === String(year)) {
      result.existingRows = g.rows.length;
      result.identicalToExisting = same;
    } else if (same) {
      result.sameAs.push({ clientId: g.clientId, clientName: names[g.clientId] || g.clientId, year: g.year });
    }
  });
  return result;
}

function getMapping_(clientId) {
  const sh = getSheetByNameSafe_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_NAMES.MAPPING);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().filter(r => String(r[0]) === String(clientId)).map(r => ({
      clientCode: String(r[1]), clientName: r[2], stdCode: String(r[3]), stdName: r[4], notes: r[5] || ''
  }));
}

/* يحدّث ربط الحسابات المرسلة فقط، ويحتفظ بربط باقي حسابات العميل (حسابات السنوات الأخرى) */
function saveMapping_(clientId, rows) {
  const sh = requireSheet_(SHEET_NAMES.MAPPING);
  const codes = {};
  (rows || []).forEach(r => codes[String(r.clientCode)] = true);
  const newRows = (rows || []).map(r => [clientId, r.clientCode, r.clientName, r.stdCode, r.stdName, r.notes || '']);
  replaceRows_(sh, r => String(r[0]) === String(clientId) && codes[String(r[1])], newRows);
  return true;
}

/* حسابات ميزان السنة المختارة + حسابات السنة السابقة (تحتاجها التدفقات النقدية كأرصدة افتتاحية) */
function mappingAccounts_(clientId, year) {
  const seen = {}, out = [];
  const add = (list, prevOnly) => list.forEach(t => {
    if (t.code && !seen[t.code]) { seen[t.code] = true; out.push({ code: t.code, name: t.name, prevOnly }); }
  });
  add(getTrialBalance_(clientId, year), false);
  add(getTrialBalance_(clientId, String(Number(year) - 1)), true);
  return out;
}

function getMappingView_(clientId, year) {
  const map = {};
  getMapping_(clientId).forEach(m => map[m.clientCode] = m);
  return mappingAccounts_(clientId, year).map(a => {
    const m = map[a.code];
    return { clientCode: a.code, clientName: a.name, stdCode: m ? m.stdCode : '', stdName: m ? m.stdName : '',
             notes: m ? m.notes : '', prevOnly: a.prevOnly };
  });
}

function autoMapClient_(clientId, year) {
  const coa = getCOA_();
  const tb  = mappingAccounts_(clientId, year);
  const existing = {};
  getMapping_(clientId).forEach(m => existing[m.clientCode] = m);

  const results = tb.map(t => {
    if (existing[t.code] && existing[t.code].stdCode) return Object.assign({}, existing[t.code], { prevOnly: t.prevOnly });
    const tName = String(t.name || '').trim();
    let exact = null, fuzzy = null;
    if (tName) {
      coa.forEach(c => {
        const cName = String(c.name || '').trim();
        if (!cName) return;
        if (!exact && cName === tName) exact = c;
        if (!fuzzy && (cName.includes(tName) || tName.includes(cName))) fuzzy = c;
      });
    }
    const best = exact || fuzzy;
    return { clientCode: t.code, clientName: t.name, stdCode: best ? best.code : '', stdName: best ? best.name : '', notes: best ? (exact ? 'مطابقة تامة' : 'مطابقة جزئية') : 'يحتاج ربط يدوي', prevOnly: t.prevOnly };
  });
  saveMapping_(clientId, results);
  return results;
}

/* ================== محرك القوائم (مبني على إعدادات دليل الحسابات) ================== */
/*
 * كل حساب في الدليل يظهر تلقائياً في القوائم حسب إعداداته — بدون أكواد ثابتة في البرنامج:
 *  - نوع الحساب   : أصول / التزامات / حقوق ملكية / إيرادات / مصروفات
 *                   (إذا كان فارغاً يُستنتج من أول رقم في الكود: 1 أصول، 2 التزامات، 3 حقوق ملكية، 4 إيرادات، 5 و6 مصروفات)
 *  - التصنيف      : متداول / غير متداول — إيراد تشغيلي / إيراد آخر — تكلفة / مصروف تشغيلي / مصروف مالي
 *  - تصنيف التدفقات: نقدية / تشغيلي / استثماري / تمويلي / غير نقدي (للإهلاك)
 *  - الترتيب      : ترتيب ظهور الحساب داخل القسم
 */
const KIND_BY_TYPE_ = { 'أصول':'asset', 'التزامات':'liability', 'خصوم':'liability', 'حقوق ملكية':'equity', 'إيرادات':'revenue', 'مصروفات':'expense' };
const KIND_BY_DIGIT_ = { '1':'asset', '2':'liability', '3':'equity', '4':'revenue', '5':'expense', '6':'expense' };
const DEBIT_KINDS_ = { asset: true, expense: true };
const BS_KINDS_ = { asset: true, liability: true, equity: true };
const GROUP_ORDER_ = { assetC:1, assetNC:2, liabilityC:3, liabilityNC:4, equity:5, revenue:6, cogs:7, opex:8, otherIncome:9, finance:10 };

function accountKind_(c) {
  const t = String(c.type || '').trim();
  if (KIND_BY_TYPE_[t]) return KIND_BY_TYPE_[t];
  const byDigit = KIND_BY_DIGIT_[String(c.code || '').charAt(0)];
  if (byDigit) return byDigit;
  if (String(c.stmt || '').includes('الدخل')) return c.normal === 'دائن' ? 'revenue' : 'expense';
  return c.normal === 'دائن' ? 'liability' : 'asset';
}

/* يحدد مكان الحساب في القوائم وتصنيفه في التدفقات */
function classifyAccount_(c) {
  const kind = accountKind_(c);
  const cls = String(c.classif || '');
  const cf = String(c.cfCat || '').trim();
  const name = String(c.name || '');
  let group;
  if (kind === 'asset' || kind === 'liability') {
    const nonCurrent = cls.includes('غير متداول') || cls.includes('طويل') ||
                       (!cls.includes('متداول') && String(c.code || '').charAt(1) !== '1');
    group = kind + (nonCurrent ? 'NC' : 'C');
  } else if (kind === 'equity') {
    group = 'equity';
  } else if (kind === 'revenue') {
    group = /آخر|أخرى|اخرى/.test(cls) ? 'otherIncome' : 'revenue';
  } else {
    group = cls.includes('تكلفة') ? 'cogs' : (/مالي|آخر|أخرى|اخرى/.test(cls) ? 'finance' : 'opex');
  }

  const isCash = kind === 'asset' && (cf === 'نقدية' ||
                 (group === 'assetC' && cf !== 'استثماري' && /نقد|بنك|بنوك|صندوق/.test(name) && !/مدين|عملاء/.test(name)));
  // حساب مقابل: مثل مجمع الإهلاك (أصل رصيده دائن)
  const isContra = DEBIT_KINDS_[kind] ? c.normal === 'دائن' : c.normal === 'مدين';
  const isNonCash = kind === 'expense' && (cf === 'غير نقدي' || /إهلاك|اهلاك|استهلاك/.test(name));

  let cfCat = '';
  if (BS_KINDS_[kind] && !isCash) {
    if (kind === 'equity') cfCat = cf === 'تمويلي' ? 'تمويلي' : '';
    else if (['تشغيلي','استثماري','تمويلي'].indexOf(cf) >= 0) cfCat = cf;
    else cfCat = { assetC:'تشغيلي', assetNC:'استثماري', liabilityC:'تشغيلي', liabilityNC:'تمويلي' }[group];
  }
  return { kind, group, isCash, isContra, isNonCash, cfCat };
}

/* رصيد كل حساب معياري = مجموع (مدين − دائن) لحسابات العميل المربوطة به */
function stdRawBalances_(tb, mapping, coaMap) {
  const tbMap = {};
  tb.forEach(t => { if (t.code) tbMap[t.code] = (tbMap[t.code] || 0) + (Number(t.debit) || 0) - (Number(t.credit) || 0); });
  const uniq = {};
  mapping.forEach(m => { if (m.clientCode && m.stdCode) uniq[m.clientCode] = m.stdCode; });
  const raw = {};
  Object.keys(uniq).forEach(cc => {
    const sc = uniq[cc];
    if (!coaMap[sc] || tbMap[cc] === undefined) return;
    raw[sc] = (raw[sc] || 0) + tbMap[cc];
  });
  return raw;
}

/* القيمة بطبيعة الحساب: الأصول والمصروفات مدينة، والباقي دائن (الحساب المقابل يظهر بالسالب) */
function signedValue_(kind, raw) { return DEBIT_KINDS_[kind] ? raw : -raw; }

function sortedAccounts_(coa) {
  return coa.map(c => Object.assign({}, c, classifyAccount_(c)))
    .sort((a, b) => GROUP_ORDER_[a.group] - GROUP_ORDER_[b.group] || (a.sort || 0) - (b.sort || 0) || String(a.code).localeCompare(String(b.code)));
}

/* كل أرقام سنة واحدة */
function buildModel_(raw, accs) {
  const val = a => signedValue_(a.kind, raw[a.code] || 0);
  const sum = f => accs.filter(f).reduce((s, a) => s + val(a), 0);
  const g = name => sum(a => a.group === name);
  const m = {
    raw, val,
    revenue: g('revenue'), cogs: g('cogs'), opex: g('opex'), otherIncome: g('otherIncome'), finance: g('finance'),
    dep: sum(a => a.isNonCash),
    curA: g('assetC'), nonCurA: g('assetNC'), curL: g('liabilityC'), nonCurL: g('liabilityNC'), eqBase: g('equity'),
    cash: sum(a => a.isCash),
    isPreClosing: accs.some(a => !BS_KINDS_[a.kind] && Math.abs(raw[a.code] || 0) > 0.0001)
  };
  m.gross = m.revenue - m.cogs;
  m.opProfit = m.gross - m.opex;
  m.netProfit = m.opProfit + m.otherIncome - m.finance;
  m.totE = m.eqBase + (m.isPreClosing ? m.netProfit : 0);
  return m;
}

/**
 * قائمة التدفقات النقدية وفق IAS 7 (الطريقة غير المباشرة + إفصاح بالطريقة المباشرة)
 * كل حساب ميزانية غير نقدي يدخل في نشاطه حسب "تصنيف التدفقات" في الدليل، لذا تتطابق النقدية دائماً.
 */
/* تنسيق رقم بفواصل الآلاف (بدون الاعتماد على Intl في Apps Script) */
function fmtN_(n) {
  const s = (Math.round(Number(n) * 100) / 100).toFixed(2);
  const parts = s.split('.');
  return parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + parts[1];
}

/* حسابات في ميزان العميل لها رصيد ولم تُربط بحساب موجود في الدليل */
function unmappedAccounts_(tb, mapping, coaMap) {
  const mapped = {};
  mapping.forEach(m => { if (m.clientCode && m.stdCode && coaMap[m.stdCode]) mapped[m.clientCode] = true; });
  const bal = {}, names = {};
  tb.forEach(t => {
    if (!t.code || mapped[t.code]) return;
    bal[t.code] = (bal[t.code] || 0) + (Number(t.debit) || 0) - (Number(t.credit) || 0);
    names[t.code] = t.name;
  });
  return Object.keys(bal).filter(c => Math.abs(bal[c]) > 0.005).map(c => ({ code: c, name: names[c], raw: bal[c] }));
}

/* فرق توازن الميزانية بعد الربط = الأصول − (الالتزامات + حقوق الملكية) */
function bsDiff_(m) {
  return Math.round(((m.curA + m.nonCurA) - (m.curL + m.nonCurL + m.totE)) * 100) / 100;
}

/* شرح سبب عدم توازن ميزانية سنة معينة */
function explainYear_(label, year, diff, unmapped) {
  if (Math.abs(diff) < 0.005) return '';
  let s = label + ' ' + year + ' غير متوازنة بعد الربط بفرق ' + fmtN_(diff);
  if (unmapped.length) {
    s += ' — بها ' + unmapped.length + ' حساب غير مربوط (مثل: ' +
         unmapped.slice(0, 5).map(u => u.code + ' ' + u.name).join('، ') + (unmapped.length > 5 ? '...' : '') + ')';
  }
  return s + '.';
}

function buildCashFlow_(accs, m, p, shown, hasPrev, prevYear, diag) {
  const d = a => m.val(a) - p.val(a);
  const bsNonCash = accs.filter(a => BS_KINDS_[a.kind] && !a.isCash && shown(a));
  const notes = [];

  // 1) التشغيلية
  // الحساب المقابل (مثل مخصص الديون) زيادته تُضاف، لذا يأخذ صياغة الالتزام
  const wc = bsNonCash.filter(a => a.cfCat === 'تشغيلي').map(a => a.kind === 'asset'
    ? { name: (a.isContra ? 'الزيادة / (النقص) في ' : '(الزيادة) / النقص في ') + a.name, value: -d(a), acc: a }
    : { name: 'الزيادة / (النقص) في ' + a.name, value: d(a), acc: a });
  const opTotal = m.netProfit + m.dep + wc.reduce((s, x) => s + x.value, 0);

  // 2) الاستثمارية
  const invItems = [];
  let contra = 0;
  bsNonCash.filter(a => a.cfCat === 'استثماري').forEach(a => {
    if (a.kind === 'asset' && a.isContra) contra += -d(a);
    else if (a.kind === 'asset') invItems.push({ name: '(شراء) / استبعاد ' + a.name, value: -d(a) });
    else invItems.push({ name: 'الزيادة / (النقص) في ' + a.name, value: d(a) });
  });
  // مجمع الإهلاك يتغير بالإهلاك المُضاف في التشغيلية؛ الباقي = استبعادات
  const adj = Math.round((contra - m.dep) * 100) / 100;
  if (adj !== 0) invItems.push({ name: 'استبعادات وتسويات أصول (صافي بعد الإهلاك)', value: adj });
  const invTotal = invItems.reduce((s, x) => s + x.value, 0);

  // 3) التمويلية
  const finItems = [];
  let explicitEquity = 0;
  bsNonCash.filter(a => a.cfCat === 'تمويلي').forEach(a => {
    if (a.kind === 'equity') { finItems.push({ name: 'زيادة / (تخفيض) ' + a.name, value: d(a) }); explicitEquity += d(a); }
    else if (a.kind === 'asset') finItems.push({ name: '(الزيادة) / النقص في ' + a.name, value: -d(a) });
    else finItems.push({ name: 'صافي التغير في ' + a.name, value: d(a) });
  });
  const otherEquity = (m.totE - p.totE) - m.netProfit - explicitEquity;   // سالب = توزيعات أرباح
  finItems.push({ name: 'توزيعات أرباح وحركات أخرى في حقوق الملكية', value: otherEquity });
  const finTotal = finItems.reduce((s, x) => s + x.value, 0);

  // 4) المطابقة
  const netChange = opTotal + invTotal + finTotal;
  const diff = Math.round((p.cash + netChange - m.cash) * 100) / 100;

  const items = [
    { section:'الأنشطة التشغيلية', items: [
        {name:'صافي ربح العام', value:m.netProfit},
        {name:'تعديلات لمطابقة صافي الربح مع صافي النقد:', heading:true},
        {name:'الإهلاك والاستهلاك', value:m.dep},
        {name:'التغيرات في رأس المال العامل:', heading:true}
      ].concat(wc.map(x => ({ name: x.name, value: x.value }))),
      totalName:'صافي النقد من الأنشطة التشغيلية', total:opTotal },
    { section:'الأنشطة الاستثمارية', items: invItems, totalName:'صافي النقد من الأنشطة الاستثمارية', total:invTotal },
    { section:'الأنشطة التمويلية', items: finItems, totalName:'صافي النقد من الأنشطة التمويلية', total:finTotal },
    { section:'صافي التغير في النقدية', isTotal:true, totalName:'صافي الزيادة / (النقص) في النقدية', total:netChange },
    { section:'النقدية وما يعادلها', items: [
        {name:'النقدية وما يعادلها أول المدة', value:p.cash},
        {name:'صافي التغير خلال العام', value:netChange}
      ].concat(diff !== 0 ? [{name:'فرق غير مطابق', value:-diff}] : []),
      totalName:'النقدية وما يعادلها آخر المدة', total:m.cash }
  ];

  // الطريقة المباشرة: المتحصلات من العملاء = الإيرادات ± التغير في حسابات العملاء والمدينين
  const arChange = wc.filter(x => x.acc.kind === 'asset' && /عملاء|مدين|ذمم|أوراق قبض/.test(x.acc.name)).reduce((s, x) => s + x.value, 0);
  const receipts = m.revenue + arChange;
  const direct = [
    { section:'الأنشطة التشغيلية (الطريقة المباشرة)', items: [
        {name:'المتحصلات النقدية من العملاء', value: receipts},
        {name:'المدفوعات النقدية للموردين والموظفين', value: opTotal - receipts - m.otherIncome + m.finance},
        {name:'متحصلات أخرى', value: m.otherIncome},
        {name:'فوائد ومصروفات مالية مدفوعة', value: -m.finance}
      ], totalName:'صافي النقد من الأنشطة التشغيلية', total:opTotal }
  ];

  if (!hasPrev) notes.push('لا يوجد ميزان مراجعة لسنة ' + prevYear + ' — تم اعتبار الأرصدة الافتتاحية صفراً (مناسب فقط للسنة الأولى للشركة). أدخل ميزان السنة السابقة للحصول على تدفقات دقيقة.');
  if (!m.isPreClosing) notes.push('ميزان المراجعة بعد الإقفال (حسابات الإيرادات والمصروفات صفر)، لذا يظهر الربح ضمن حركة حقوق الملكية وليس في الأنشطة التشغيلية.');
  if (!accs.some(a => a.isCash)) notes.push('لا يوجد حساب نقدية في الدليل. حدد "نقدية" في عمود تصنيف التدفقات لحسابات النقدية والبنوك.');
  if (diff !== 0) {
    // رياضياً: فرق النقدية = فرق توازن الميزانية الافتتاحية − فرق توازن الميزانية الختامية
    const why = [
      explainYear_('ميزانية', diag.year, diag.curDiff, diag.curUnmapped),
      explainYear_('الأرصدة الافتتاحية (ميزان', prevYear + ')', diag.prevDiff, diag.prevUnmapped)
    ].filter(Boolean);
    let msg = 'يوجد فرق غير مطابق في النقدية = ' + fmtN_(diff) + '. ';
    msg += why.length ? 'السبب: ' + why.join(' ') : 'السبب غالباً حسابات غير مربوطة أو ميزان غير متوازن.';
    if (diag.prevUnmapped.length) msg += ' الحل: افتح "ربط الحسابات" لسنة ' + diag.year + ' (تعرض الآن حسابات ' + prevYear + ' أيضاً) واربطها ثم أعد إعداد القوائم.';
    else if (diag.curUnmapped.length) msg += ' الحل: اربط هذه الحسابات من صفحة "ربط الحسابات" ثم أعد إعداد القوائم.';
    notes.push(msg);
  }
  notes.push('تم تصنيف الفوائد المدفوعة ضمن الأنشطة التشغيلية (IAS 7 فقرة 33).');

  return { items, direct, notes, opTotal };
}

/**
 * ورقة العمل: كل حساب في ميزان العميل وأين ظهر في القوائم.
 * المبلغ بنفس إشارة القائمة، فمجموع حسابات كل بند = رقمه في القائمة.
 */
function buildWorkingPaper_(tb, mapping, coa) {
  const coaMap = {}; coa.forEach(c => coaMap[c.code] = c);
  const mapBy = {}; mapping.forEach(m => { if (m.clientCode) mapBy[m.clientCode] = m; });

  const tbMap = {}, codes = [];
  tb.forEach(t => {
    if (!t.code) return;
    if (!tbMap[t.code]) { tbMap[t.code] = { name: t.name, raw: 0 }; codes.push(t.code); }
    tbMap[t.code].raw += (Number(t.debit) || 0) - (Number(t.credit) || 0);
  });

  const rows = codes.map(code => {
    const t = tbMap[code], m = mapBy[code];
    const std = (m && m.stdCode) ? coaMap[m.stdCode] : null;
    if (!std) {
      return { stmt: 'غير مربوط', line: 'حسابات غير مربوطة (لم تدخل القوائم)', code, name: t.name,
               stdCode: '', stdName: '', amount: t.raw, status: 'unmapped', key: [99, 0, '', code] };
    }
    const c = classifyAccount_(std);
    const isBS = BS_KINDS_[c.kind];
    return { stmt: isBS ? 'الميزانية العمومية' : 'قائمة الدخل', line: std.name, code, name: t.name,
             stdCode: std.code, stdName: std.name,
             amount: isBS ? signedValue_(c.kind, t.raw) : -t.raw,   // الدخل: الإيراد موجب والمصروف سالب
             status: 'ok', key: [GROUP_ORDER_[c.group], Number(std.sort) || 0, String(std.code), code] };
  });

  rows.sort((a, b) => {
    for (let i = 0; i < 4; i++) {
      if (a.key[i] < b.key[i]) return -1;
      if (a.key[i] > b.key[i]) return 1;
    }
    return 0;
  });
  rows.forEach(r => delete r.key);
  return rows;
}

function getWorkingPaper_(clientId, year) {
  const tb = getTrialBalance_(clientId, year);
  if (!tb.length) throw new Error('لا يوجد ميزان مراجعة لهذا العميل والسنة');
  return buildWorkingPaper_(tb, getMapping_(clientId), getCOA_());
}

function computeStatements_(clientId, year) {
  const tb = getTrialBalance_(clientId, year);
  if (!tb.length) throw new Error('لا يوجد ميزان مراجعة لهذا العميل والسنة');
  const mapping = getMapping_(clientId);
  if (!mapping.length) throw new Error('يرجى ربط حسابات الشركة بدليل المكتب أولاً');

  const coa = getCOA_();
  const coaMap = {}; coa.forEach(c => coaMap[c.code] = c);
  const accs = sortedAccounts_(coa);

  // السنة الحالية والسابقة (السابقة للتدفقات النقدية)
  const prevYear = String(Number(year) - 1);
  const tbPrev = getTrialBalance_(clientId, prevYear);
  const m = buildModel_(stdRawBalances_(tb, mapping, coaMap), accs);
  const p = buildModel_(tbPrev.length ? stdRawBalances_(tbPrev, mapping, coaMap) : {}, accs);

  // يظهر الحساب إذا كان له رصيد في السنة الحالية أو السابقة
  const shown = a => Math.abs(m.raw[a.code] || 0) > 0.0001 || Math.abs(p.raw[a.code] || 0) > 0.0001;
  const lines = group => accs.filter(a => a.group === group && shown(a));

  /* ---------- قائمة الدخل ---------- */
  const isItems = [];
  const addGroup = (group, heading, sign) => {
    const list = lines(group);
    if (!list.length) return;
    isItems.push({ name: heading, heading: true });
    list.forEach(a => isItems.push({ name: a.name, value: sign * m.val(a) }));
  };
  addGroup('revenue', 'الإيرادات:', 1);
  addGroup('cogs', 'يخصم: تكلفة المبيعات:', -1);
  isItems.push({ name: 'مجمل الربح', value: m.gross, isTotal: true });
  addGroup('opex', 'يخصم: المصروفات التشغيلية:', -1);
  isItems.push({ name: 'الربح التشغيلي', value: m.opProfit, isTotal: true });
  addGroup('otherIncome', 'يضاف: إيرادات أخرى:', 1);
  addGroup('finance', 'يخصم: مصروفات مالية وأخرى:', -1);
  isItems.push({ name: 'صافي الربح', value: m.netProfit, isTotal: true, isFinal: true });
  const isObj = { items: isItems };

  /* ---------- الميزانية العمومية ---------- */
  const items = group => lines(group).map(a => ({ name: a.name, value: m.val(a) }));
  const totA = m.curA + m.nonCurA;
  const totLE = m.curL + m.nonCurL + m.totE;
  const bsObj = {
    sections: [
      { section: 'الأصول المتداولة', items: items('assetC'), totalName: 'إجمالي الأصول المتداولة', total: m.curA },
      { section: 'الأصول غير المتداولة', items: items('assetNC'), totalName: 'إجمالي الأصول غير المتداولة', total: m.nonCurA },
      { section: 'إجمالي الأصول', isTotal: true, totalName: 'إجمالي الأصول', total: totA },
      { section: 'الالتزامات المتداولة', items: items('liabilityC'), totalName: 'إجمالي الالتزامات المتداولة', total: m.curL },
      { section: 'الالتزامات غير المتداولة', items: items('liabilityNC'), totalName: 'إجمالي الالتزامات غير المتداولة', total: m.nonCurL },
      { section: 'حقوق الملكية', items: items('equity').concat(m.isPreClosing ? [{ name: 'صافي ربح العام', value: m.netProfit }] : []),
        totalName: 'إجمالي حقوق الملكية', total: m.totE },
      { section: 'إجمالي الالتزامات وحقوق الملكية', isTotal: true, totalName: 'إجمالي الالتزامات وحقوق الملكية', total: totLE }
    ],
    notes: []
  };
  const diag = {
    year: String(year),
    curDiff: bsDiff_(m), curUnmapped: unmappedAccounts_(tb, mapping, coaMap),
    prevDiff: tbPrev.length ? bsDiff_(p) : 0, prevUnmapped: unmappedAccounts_(tbPrev, mapping, coaMap)
  };
  if (diag.curDiff !== 0) {
    bsObj.notes.push(explainYear_('ميزانية', diag.year, diag.curDiff, diag.curUnmapped) +
      ' راجع ورقة العمل وصفحة "ربط الحسابات".');
  }

  /* ---------- التدفقات والمؤشرات ---------- */
  const cf = buildCashFlow_(accs, m, p, shown, tbPrev.length > 0, prevYear, diag);
  const cfObj = { items: cf.items, direct: cf.direct, notes: cf.notes };

  const safe = (a, b) => b ? a / b : 0;
  const ratios = [
    {cat:'السيولة', name:'نسبة التداول', value:safe(m.curA, m.curL), eval: safe(m.curA, m.curL)>=1.5?'جيد':'ضعيف'},
    {cat:'الربحية', name:'هامش الربح الإجمالي %', value:safe(m.gross, m.revenue)*100, eval: safe(m.gross, m.revenue)>=0.25?'جيد':'ضعيف'},
    {cat:'الربحية', name:'هامش صافي الربح %', value:safe(m.netProfit, m.revenue)*100, eval: safe(m.netProfit, m.revenue)>=0.1?'جيد':'ضعيف'},
    {cat:'التدفقات النقدية', name:'نسبة التدفق النقدي التشغيلي', value:safe(cf.opTotal, m.curL), eval: safe(cf.opTotal, m.curL)>=1?'جيد':'ضعيف'},
    {cat:'التدفقات النقدية', name:'جودة الأرباح (النقد التشغيلي / صافي الربح)', value:safe(cf.opTotal, m.netProfit), eval: (m.netProfit>0 && safe(cf.opTotal, m.netProfit)>=1)?'جيد':'ضعيف'}
  ];

  const wp = buildWorkingPaper_(tb, mapping, coa);
  saveStatementsToSheets_(clientId, year, bsObj, isObj, cfObj, ratios, wp);
  return { bs: bsObj, is: isObj, cf: cfObj, ratios: ratios };
}

/* يضيف عنوان عمود "النوع" للملفات القديمة */
function ensureHeaderCell_(sh, col, title) {
  if (sh && !sh.getRange(1, col).getValue()) {
    sh.getRange(1, col).setValue(title).setFontWeight('bold').setBackground(HEADER_BG).setFontColor(HEADER_FG);
  }
  return sh;
}

function saveStatementsToSheets_(clientId, year, bsObj, isObj, cfObj, ratios, wp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const match = r => isClientYear_(r, clientId, year);

  // ورقة العمل (تُنشأ تلقائياً للملفات القديمة)
  const wpRows = (wp || []).map(r => [clientId, year, r.stmt, r.line, r.code, r.name, r.stdCode ? r.stdCode + ' - ' + r.stdName : '', r.amount]);
  replaceRows_(ensureSheet_(ss, SHEET_NAMES.WP, WP_HEADERS), match, wpRows);

  // عمود "النوع" في كل القوائم: بند / عنوان / إجمالي / نهائي / ملاحظة
  const bsRows = [];
  bsObj.sections.forEach(sec => {
    (sec.items || []).forEach(it => bsRows.push([clientId, year, sec.section, it.name, it.value, 'بند']));
    if (sec.total !== undefined) bsRows.push([clientId, year, sec.section, sec.totalName || 'الإجمالي', sec.total, 'إجمالي']);
  });
  (bsObj.notes || []).forEach(n => bsRows.push([clientId, year, 'ملاحظات', n, '', 'ملاحظة']));
  replaceRows_(ensureHeaderCell_(getSheetByNameSafe_(ss, SHEET_NAMES.BS), 6, 'النوع'), match, bsRows);

  const isRows = isObj.items.map(it => [clientId, year, it.name, it.heading ? '' : it.value,
    it.heading ? 'عنوان' : (it.isFinal ? 'نهائي' : (it.isTotal ? 'إجمالي' : 'بند'))]);
  replaceRows_(ensureHeaderCell_(getSheetByNameSafe_(ss, SHEET_NAMES.IS), 5, 'النوع'), match, isRows);

  const cfRows = [];
  (cfObj.items || []).concat(cfObj.direct || []).forEach(sec => {
    (sec.items || []).forEach(it => cfRows.push([clientId, year, sec.section, it.name, it.heading ? '' : it.value, it.heading ? 'عنوان' : 'بند']));
    if (sec.total !== undefined) cfRows.push([clientId, year, sec.section, sec.totalName || ('إجمالي ' + sec.section), sec.total, 'إجمالي']);
  });
  (cfObj.notes || []).forEach(n => cfRows.push([clientId, year, 'ملاحظات', n, '', 'ملاحظة']));
  replaceRows_(ensureHeaderCell_(getSheetByNameSafe_(ss, SHEET_NAMES.CF), 6, 'النوع'), match, cfRows);

  const ratioRows = ratios.map(r => [clientId, year, r.cat + ' - ' + r.name, r.value, r.eval]);
  replaceRows_(getSheetByNameSafe_(ss, SHEET_NAMES.RATIOS), match, ratioRows);
}

function getSavedStatements_(clientId, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fetch = (sheetName) => {
    const sh = getSheetByNameSafe_(ss, sheetName);
    if (!sh) return [];
    return sh.getDataRange().getValues().filter(r => isClientYear_(r, clientId, year));
  };

  const bsRows = fetch(SHEET_NAMES.BS);
  const isRows = fetch(SHEET_NAMES.IS);
  const cfRows = fetch(SHEET_NAMES.CF);
  const ratRows = fetch(SHEET_NAMES.RATIOS);

  if (!bsRows.length && !isRows.length && !cfRows.length) {
    throw new Error('لا توجد قوائم محفوظة لهذا العميل وهذه السنة. يرجى النقر على "إعداد القوائم" أولاً.');
  }

  // الصفوف القديمة (بدون عمود النوع) تُميَّز بالاسم
  const bsSections = [], bsNotes = [];
  bsRows.forEach(r => {
    const secName = String(r[2]), itemName = String(r[3]), val = r[4], kind = String(r[5] || '');
    if (kind === 'ملاحظة') { bsNotes.push(itemName); return; }
    let sec = bsSections.find(s => s.section === secName);
    if (!sec) { sec = { section: secName }; bsSections.push(sec); }
    const isTot = kind ? kind === 'إجمالي'
                       : (itemName === secName || itemName.includes('إجمالي') || secName.startsWith('إجمالي'));
    if (isTot) {
      sec.total = val; sec.totalName = itemName;
      if (secName.startsWith('إجمالي')) sec.isTotal = true;
    } else {
      if (!sec.items) sec.items = [];
      sec.items.push({ name: itemName, value: val });
    }
  });

  const isItems = isRows.map(r => {
    const name = String(r[2]), kind = String(r[4] || '');
    if (kind === 'عنوان') return { name: name, heading: true };
    if (kind) return { name: name, value: r[3], isTotal: kind === 'إجمالي' || kind === 'نهائي', isFinal: kind === 'نهائي' };
    return { name: name, value: r[3], isTotal: name.includes('مجمل الربح') || name.includes('الربح التشغيلي'), isFinal: name.includes('صافي الربح') };
  });

  const cfSections = [], cfDirect = [], cfNotes = [];
  cfRows.forEach(r => {
    const secName = String(r[2]), itemName = String(r[3]), val = r[4], kind = String(r[5] || '');
    if (kind === 'ملاحظة') { cfNotes.push(itemName); return; }
    const list = secName.includes('الطريقة المباشرة') ? cfDirect : cfSections;
    let sec = list.find(s => s.section === secName);
    if (!sec) { sec = { section: secName }; list.push(sec); }
    const isTotalRow = kind ? kind === 'إجمالي' : (itemName.startsWith('إجمالي') || secName === 'صافي التغير في النقدية');
    if (isTotalRow) {
      sec.total = val; sec.totalName = itemName;
      if (secName === 'صافي التغير في النقدية') sec.isTotal = true;
    } else {
      if (!sec.items) sec.items = [];
      sec.items.push(kind === 'عنوان' ? { name: itemName, heading: true } : { name: itemName, value: val });
    }
  });

  const ratios = ratRows.map(r => {
    const full = String(r[2]);
    return { cat: full.split(' - ')[0], name: full.split(' - ')[1] || full, value: r[3], eval: r[4] };
  });

  return { bs: { sections: bsSections, notes: bsNotes }, is: { items: isItems },
           cf: { items: cfSections, direct: cfDirect, notes: cfNotes }, ratios: ratios };
}

/* ================== مقارنة السنوات ================== */
/* يدمج قائمتين مرتبتين مع الحفاظ على الترتيب (البند الموجود في سنة واحدة فقط لا يضيع) */
function mergeOrdered_(a, b) {
  const out = a.slice();
  let last = -1;
  b.forEach(x => {
    const i = out.indexOf(x);
    if (i >= 0) { last = i; return; }
    out.splice(last + 1, 0, x);
    last++;
  });
  return out;
}

function getComparisonData_(clientId, year1, year2) {
  let y1Data = null, y2Data = null;
  try { y1Data = getSavedStatements_(clientId, year1); } catch(e) {}
  try { y2Data = getSavedStatements_(clientId, year2); } catch(e) {}

  if (!y1Data && !y2Data) throw new Error('لا توجد بيانات محفوظة لكلا السنتين. يرجى إعداد القوائم أولاً.');
  if (!y1Data) throw new Error('لا توجد قوائم محفوظة لسنة ' + year1);
  if (!y2Data) throw new Error('لا توجد قوائم محفوظة لسنة ' + year2);

  const s1 = y1Data.bs.sections, s2 = y2Data.bs.sections;
  const find = (list, name) => list.find(s => s.section === name) || {};
  const val = (sec, name) => { const it = (sec.items || []).find(i => i.name === name); return it ? it.value : 0; };

  const bsItems = [];
  mergeOrdered_(s1.map(s => s.section), s2.map(s => s.section)).forEach(name => {
    const a = find(s1, name), b = find(s2, name);
    mergeOrdered_((a.items || []).map(i => i.name), (b.items || []).map(i => i.name))
      .forEach(n => bsItems.push({ section: name, name: n, y1: val(a, n), y2: val(b, n), isTotal: false }));
    if (a.total !== undefined || b.total !== undefined) {
      bsItems.push({ section: name, name: a.totalName || b.totalName || 'إجمالي', y1: a.total || 0, y2: b.total || 0, isTotal: true });
    }
  });

  const i1 = y1Data.is.items.filter(i => !i.heading), i2 = y2Data.is.items.filter(i => !i.heading);
  const isItems = mergeOrdered_(i1.map(i => i.name), i2.map(i => i.name)).map(n => {
    const a = i1.find(i => i.name === n), b = i2.find(i => i.name === n), ref = a || b;
    return { name: n, y1: a ? a.value : 0, y2: b ? b.value : 0, isTotal: ref.isTotal, isFinal: ref.isFinal };
  });

  return { bs: bsItems, is: isItems };
}
