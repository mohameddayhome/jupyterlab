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
  const user = getSession_(token);
  if (!user) throw new Error('SESSION_EXPIRED');
  if (fn === 'saveSettings' && user.role !== 'مدير') throw new Error('تعديل الإعدادات متاح للمدير فقط');
  const routes = {
    getCOA: getCOA_, saveCOAAccount: saveCOAAccount_, deleteCOAAccount: deleteCOAAccount_,
    getDashboardStats: getDashboardStats_,
    getClients: getClients_, saveClient: saveClient_, deleteClient: deleteClient_,
    getTrialBalance: getTrialBalance_, saveTrialBalance: saveTrialBalance_,
    getMapping: getMapping_, getMappingView: getMappingView_, saveMapping: saveMapping_, autoMapClient: autoMapClient_,
    computeStatements: computeStatements_, getSavedStatements: getSavedStatements_,
    getComparisonData: getComparisonData_,
    getWorkingPaper: getWorkingPaper_, checkTrialBalanceDuplicate: checkTrialBalanceDuplicate_,
    getNotes: getNotes_, getSettings: getSettings_, saveSettings: saveSettings_
  };
  const f = routes[fn];
  if (!f) throw new Error('عملية غير مسموحة: ' + fn);
  return f.apply(null, args || []);
}

/* ================== إعدادات المكتب (الاسم والشعار) ================== */
// تُحفظ في خصائص السكربت؛ الشعار مقسّم لأجزاء لأن حد الخاصية الواحدة 9KB
const DEFAULT_OFFICE_NAME = 'مكتب المدقق';
const LOGO_CHUNK = 8000;

function getSettings_() {
  const props = PropertiesService.getScriptProperties();
  const n = Number(props.getProperty('logo_chunks') || 0);
  let logo = '';
  for (let i = 0; i < n; i++) logo += props.getProperty('logo_' + i) || '';
  return {
    officeName: props.getProperty('officeName') || DEFAULT_OFFICE_NAME,
    officeSubtitle: props.getProperty('officeSubtitle') || '',
    logo: logo
  };
}

/* s.logo: data URL لصورة، أو '' لحذف الشعار، أو undefined لعدم تغييره */
function saveSettings_(s) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('officeName', String(s.officeName || '').trim() || DEFAULT_OFFICE_NAME);
  props.setProperty('officeSubtitle', String(s.officeSubtitle || '').trim());
  if (s.logo !== undefined && s.logo !== null) {
    const logo = String(s.logo);
    if (logo && !/^data:image\/(png|jpeg|gif|webp);base64,[A-Za-z0-9+\/=]+$/.test(logo)) throw new Error('صيغة الشعار غير مدعومة (PNG أو JPG)');
    if (logo.length > 300000) throw new Error('حجم الشعار كبير جداً — استخدم صورة أصغر');
    const old = Number(props.getProperty('logo_chunks') || 0);
    for (let i = 0; i < old; i++) props.deleteProperty('logo_' + i);
    const parts = {};
    const n = Math.ceil(logo.length / LOGO_CHUNK);
    for (let i = 0; i < n; i++) parts['logo_' + i] = logo.substr(i * LOGO_CHUNK, LOGO_CHUNK);
    parts.logo_chunks = String(n);
    props.setProperties(parts);
  }
  return getSettings_();
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
  // حسابات يحتاجها التحليل المالي (تُعرف من الاسم)
  const isInventory  = kind === 'asset' && /مخزون|بضاع/.test(name);
  const isReceivable = kind === 'asset' && !isCash && (/عملاء|مدين|ذمم مدينة|أوراق قبض/.test(name) || /مخصص.*(ديون|مشكوك)/.test(name));
  const isPayable    = kind === 'liability' && /مورد|دائن|ذمم دائنة|أوراق دفع/.test(name);
  const isCapital    = kind === 'equity' && /رأس المال|راس المال|رأس مال|راس مال/.test(name);
  return { kind, group, isCash, isContra, isNonCash, cfCat, isInventory, isReceivable, isPayable, isCapital };
}

/* رصيد كل حساب معياري = مجموع (مدين − دائن) لحسابات العميل المربوطة به */
/* حساب العميل ← الحساب المعياري (آخر ربط للحساب، ويُتجاهل الربط بحساب محذوف من الدليل) */
function mappedStd_(mapping, coaMap) {
  const uniq = {};
  mapping.forEach(m => { if (m.clientCode && m.stdCode) uniq[m.clientCode] = m.stdCode; });
  Object.keys(uniq).forEach(cc => { if (!coaMap[uniq[cc]]) delete uniq[cc]; });
  return uniq;
}

function stdRawBalances_(tb, mapping, coaMap) {
  const tbMap = {};
  tb.forEach(t => { if (t.code) tbMap[t.code] = (tbMap[t.code] || 0) + (Number(t.debit) || 0) - (Number(t.credit) || 0); });
  const uniq = mappedStd_(mapping, coaMap);
  const raw = {};
  Object.keys(uniq).forEach(cc => {
    if (tbMap[cc] === undefined) return;
    raw[uniq[cc]] = (raw[uniq[cc]] || 0) + tbMap[cc];
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
    inventory: sum(a => a.isInventory), receivables: sum(a => a.isReceivable),
    payables: sum(a => a.isPayable), capital: sum(a => a.isCapital),
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

/* بيانات سنة واحدة + السنة السابقة، مشتركة بين القوائم والمرفقات */
function loadYearData_(clientId, year) {
  const tb = getTrialBalance_(clientId, year);
  if (!tb.length) throw new Error('لا يوجد ميزان مراجعة لهذا العميل والسنة');
  const mapping = getMapping_(clientId);
  if (!mapping.length) throw new Error('يرجى ربط حسابات الشركة بدليل المكتب أولاً');

  const coa = getCOA_();
  const coaMap = {}; coa.forEach(c => coaMap[c.code] = c);
  const accs = sortedAccounts_(coa);

  // السنة الحالية والسابقة (السابقة للتدفقات النقدية وأرقام المقارنة)
  const prevYear = String(Number(year) - 1);
  const tbPrev = getTrialBalance_(clientId, prevYear);
  const m = buildModel_(stdRawBalances_(tb, mapping, coaMap), accs);
  const p = buildModel_(tbPrev.length ? stdRawBalances_(tbPrev, mapping, coaMap) : {}, accs);

  // يظهر الحساب إذا كان له رصيد في السنة الحالية أو السابقة
  const shown = a => Math.abs(m.raw[a.code] || 0) > 0.0001 || Math.abs(p.raw[a.code] || 0) > 0.0001;
  return { tb, tbPrev, mapping, coa, coaMap, accs, prevYear, m, p, shown };
}

const GROUP_TITLES_ = {
  assetC:'الأصول المتداولة', assetNC:'الأصول غير المتداولة', liabilityC:'الالتزامات المتداولة',
  liabilityNC:'الالتزامات غير المتداولة', equity:'حقوق الملكية', revenue:'الإيرادات', cogs:'تكلفة المبيعات',
  opex:'المصروفات التشغيلية', otherIncome:'إيرادات أخرى', finance:'مصروفات مالية وأخرى'
};

/**
 * المرفقات (الإيضاحات المتممة): لكل بند في القوائم إيضاح مرقّم يفصّل حسابات العميل المكوّنة له
 * للسنة الحالية والسابقة. الترقيم بنفس ترتيب ظهور البنود في الميزانية ثم قائمة الدخل.
 * المبالغ بطبيعة الحساب (المصروفات موجبة كما هو متعارف عليه في الإيضاحات).
 */
function buildNotes_(d) {
  // الإيضاحات تفصّل أرصدة سنة التقرير فقط؛ البند الذي ليس له رصيد هذه السنة لا يأخذ إيضاحاً
  const toStd = mappedStd_(d.mapping, d.coaMap);
  const byStd = {};
  d.tb.forEach(t => {
    const sc = toStd[t.code];
    if (!t.code || !sc) return;
    const bucket = byStd[sc] = byStd[sc] || {};
    const e = bucket[t.code] = bucket[t.code] || { code: t.code, name: t.name, raw: 0 };
    e.raw += (Number(t.debit) || 0) - (Number(t.credit) || 0);
  });

  const notes = [];
  d.accs.filter(a => Math.abs(d.m.raw[a.code] || 0) > 0.0001).forEach(a => {
    const rows = Object.keys(byStd[a.code] || {}).map(k => byStd[a.code][k])
      .map(r => ({ code: r.code, name: r.name, amount: signedValue_(a.kind, r.raw) }))
      .filter(r => Math.abs(r.amount) > 0.0001)
      .sort((x, y) => String(x.code).localeCompare(String(y.code)));
    notes.push({
      no: notes.length + 1, stdCode: a.code, title: a.name,
      stmt: BS_KINDS_[a.kind] ? 'الميزانية العمومية' : 'قائمة الدخل', group: GROUP_TITLES_[a.group],
      rows: rows, total: d.m.val(a)
    });
  });
  return notes;
}

function getNotes_(clientId, year) {
  const d = loadYearData_(clientId, year);
  const client = getClients_().find(c => String(c.id) === String(clientId)) || {};
  return { year: String(year),
           clientName: client.name || '', currency: client.currency || '', endDate: client.endDate || '',
           notes: buildNotes_(d) };
}

function computeStatements_(clientId, year) {
  const d = loadYearData_(clientId, year);
  const { tb, tbPrev, mapping, coa, coaMap, accs, prevYear, m, p, shown } = d;
  const lines = group => accs.filter(a => a.group === group && shown(a));
  // رقم الإيضاح لكل بند
  const noteNo = {};
  buildNotes_(d).forEach(n => noteNo[n.stdCode] = n.no);

  // أرقام المقارنة للسنة السابقة (IAS 1 فقرة 38) — null إذا لم يوجد ميزان للسنة السابقة
  const hasPrev = tbPrev.length > 0;
  const pv = v => hasPrev ? v : null;

  /* ---------- قائمة الدخل ---------- */
  const isItems = [];
  const addGroup = (group, heading, sign) => {
    const list = lines(group);
    if (!list.length) return;
    isItems.push({ name: heading, heading: true });
    list.forEach(a => isItems.push({ name: a.name, value: sign * m.val(a), prev: pv(sign * p.val(a)), note: noteNo[a.code] }));
  };
  addGroup('revenue', 'الإيرادات:', 1);
  addGroup('cogs', 'يخصم: تكلفة المبيعات:', -1);
  isItems.push({ name: 'مجمل الربح', value: m.gross, prev: pv(p.gross), isTotal: true });
  addGroup('opex', 'يخصم: المصروفات التشغيلية:', -1);
  isItems.push({ name: 'الربح التشغيلي', value: m.opProfit, prev: pv(p.opProfit), isTotal: true });
  addGroup('otherIncome', 'يضاف: إيرادات أخرى:', 1);
  addGroup('finance', 'يخصم: مصروفات مالية وأخرى:', -1);
  isItems.push({ name: 'صافي الربح', value: m.netProfit, prev: pv(p.netProfit), isTotal: true, isFinal: true });
  const isObj = { items: isItems };

  /* ---------- الميزانية العمومية ---------- */
  const items = group => lines(group).map(a => ({ name: a.name, value: m.val(a), prev: pv(p.val(a)), note: noteNo[a.code] }));
  const npLine = (m.isPreClosing || (hasPrev && p.isPreClosing))
    ? [{ name: 'صافي ربح العام', value: m.isPreClosing ? m.netProfit : 0, prev: pv(p.isPreClosing ? p.netProfit : 0) }] : [];
  const bsObj = {
    sections: [
      { section: 'الأصول المتداولة', items: items('assetC'), totalName: 'إجمالي الأصول المتداولة', total: m.curA, prevTotal: pv(p.curA) },
      { section: 'الأصول غير المتداولة', items: items('assetNC'), totalName: 'إجمالي الأصول غير المتداولة', total: m.nonCurA, prevTotal: pv(p.nonCurA) },
      { section: 'إجمالي الأصول', isTotal: true, totalName: 'إجمالي الأصول', total: m.curA + m.nonCurA, prevTotal: pv(p.curA + p.nonCurA) },
      { section: 'الالتزامات المتداولة', items: items('liabilityC'), totalName: 'إجمالي الالتزامات المتداولة', total: m.curL, prevTotal: pv(p.curL) },
      { section: 'الالتزامات غير المتداولة', items: items('liabilityNC'), totalName: 'إجمالي الالتزامات غير المتداولة', total: m.nonCurL, prevTotal: pv(p.nonCurL) },
      { section: 'حقوق الملكية', items: items('equity').concat(npLine), totalName: 'إجمالي حقوق الملكية', total: m.totE, prevTotal: pv(p.totE) },
      { section: 'إجمالي الالتزامات وحقوق الملكية', isTotal: true, totalName: 'إجمالي الالتزامات وحقوق الملكية',
        total: m.curL + m.nonCurL + m.totE, prevTotal: pv(p.curL + p.nonCurL + p.totE) }
    ],
    notes: []
  };
  if (!hasPrev) bsObj.notes.push('لا يوجد ميزان مراجعة لسنة ' + prevYear + ' — لن تظهر أرقام المقارنة. أدخل ميزان السنة السابقة لعرض السنتين كما يتطلب IAS 1.');
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
  const prev2Year = String(Number(year) - 2);
  const tbPrev2 = hasPrev ? getTrialBalance_(clientId, prev2Year) : [];
  const pp = buildModel_(tbPrev2.length ? stdRawBalances_(tbPrev2, mapping, coaMap) : {}, accs);
  // نفس البنود في السنتين حتى تتطابق الأسطر
  const shownCF = a => shown(a) || Math.abs(pp.raw[a.code] || 0) > 0.0001;
  const cf = buildCashFlow_(accs, m, p, shownCF, hasPrev, prevYear, diag);
  let cfPrev = null;
  if (hasPrev && tbPrev2.length) {
    cfPrev = buildCashFlow_(accs, p, pp, shownCF, true, prev2Year, {
      year: prevYear, curDiff: bsDiff_(p), curUnmapped: unmappedAccounts_(tbPrev, mapping, coaMap),
      prevDiff: bsDiff_(pp), prevUnmapped: unmappedAccounts_(tbPrev2, mapping, coaMap) });
  }
  const cfObj = { items: mergeComparative_(cf.items, cfPrev && cfPrev.items),
                  direct: mergeComparative_(cf.direct, cfPrev && cfPrev.direct), notes: cf.notes.slice() };
  if (hasPrev && !tbPrev2.length) {
    cfObj.notes.push('أرقام المقارنة لقائمة التدفقات النقدية (سنة ' + prevYear + ') تحتاج ميزان مراجعة سنة ' + prev2Year + ' كأرصدة افتتاحية — أدخله لعرض السنتين.');
  }
  if (cfPrev) cfPrev.notes.filter(n => n.indexOf('فرق') >= 0).forEach(n => cfObj.notes.push('أرقام المقارنة (' + prevYear + '): ' + n));

  // المؤشرات للسنتين: الأرصدة المتوسطة تستخدم السنة السابقة لها إن وُجدت
  const ratios = buildRatios_(m, hasPrev ? p : null, cf.opTotal,
                              hasPrev ? p : null, (hasPrev && tbPrev2.length) ? pp : null, cfPrev ? cfPrev.opTotal : null);

  const wp = buildWorkingPaper_(tb, mapping, coa);
  saveStatementsToSheets_(clientId, year, bsObj, isObj, cfObj, ratios, wp);
  return { bs: bsObj, is: isObj, cf: cfObj, ratios: ratios, years: { cur: String(year), prev: prevYear } };
}

/* ================== المؤشرات المالية (التحليل المالي) ================== */
/*
 * كل مؤشر: المعادلة، الوحدة، قيمة سنة التقرير والسنة السابقة، ومعيار تقييم عام.
 * المعايير استرشادية وتختلف حسب نشاط المنشأة.
 * الأرصدة "المتوسطة" = (أول المدة + آخر المدة) ÷ 2، وإذا لم تتوفر السنة السابقة يُستخدم رصيد آخر المدة.
 */
const RATIO_DEFS_ = [
  // [التصنيف، الاسم، المعادلة، الوحدة، الدالة، [جيد، مقبول]، أعلى أفضل؟]
  ['السيولة', 'نسبة التداول', 'الأصول المتداولة ÷ الالتزامات المتداولة', 'مرة', r => r.div(r.c.curA, r.c.curL), [2, 1], true],
  ['السيولة', 'نسبة السيولة السريعة', '(الأصول المتداولة − المخزون) ÷ الالتزامات المتداولة', 'مرة', r => r.div(r.c.curA - r.c.inventory, r.c.curL), [1, 0.7], true],
  ['السيولة', 'نسبة النقدية', 'النقدية ÷ الالتزامات المتداولة', 'مرة', r => r.div(r.c.cash, r.c.curL), [0.5, 0.2], true],
  ['السيولة', 'رأس المال العامل', 'الأصول المتداولة − الالتزامات المتداولة', 'مبلغ', r => r.c.curA - r.c.curL, [0.0001, 0], true],

  ['الربحية', 'هامش مجمل الربح', 'مجمل الربح ÷ المبيعات', '%', r => r.pct(r.c.gross, r.c.revenue), [40, 20], true],
  ['الربحية', 'هامش الربح التشغيلي', 'الربح التشغيلي ÷ المبيعات', '%', r => r.pct(r.c.opProfit, r.c.revenue), [15, 5], true],
  ['الربحية', 'العائد على المبيعات (هامش صافي الربح)', 'صافي الربح ÷ المبيعات', '%', r => r.pct(r.c.netProfit, r.c.revenue), [10, 5], true],
  ['الربحية', 'العائد على الأصول (ROA)', 'صافي الربح ÷ متوسط إجمالي الأصول', '%', r => r.pct(r.c.netProfit, r.avg(x => x.curA + x.nonCurA)), [10, 5], true],
  ['الربحية', 'العائد على حقوق الملكية (ROE)', 'صافي الربح ÷ متوسط حقوق الملكية', '%', r => r.pct(r.c.netProfit, r.avg(x => x.totE)), [15, 8], true],
  ['الربحية', 'العائد على رأس المال المدفوع', 'صافي الربح ÷ رأس المال المدفوع', '%', r => r.pct(r.c.netProfit, r.c.capital), [20, 10], true],
  ['الربحية', 'العائد على رأس المال المستخدم (ROCE)', 'الربح قبل الفوائد ÷ (إجمالي الأصول − الالتزامات المتداولة)', '%', r => r.pct(r.c.netProfit + r.c.finance, r.c.curA + r.c.nonCurA - r.c.curL), [15, 8], true],

  ['النشاط والكفاءة', 'معدل دوران الأصول', 'المبيعات ÷ متوسط إجمالي الأصول', 'مرة', r => r.div(r.c.revenue, r.avg(x => x.curA + x.nonCurA)), [1, 0.5], true],
  ['النشاط والكفاءة', 'معدل دوران المدينين', 'المبيعات ÷ متوسط المدينين', 'مرة', r => r.div(r.c.revenue, r.avg(x => x.receivables)), [8, 4], true],
  ['النشاط والكفاءة', 'متوسط فترة التحصيل', '365 ÷ معدل دوران المدينين', 'يوم', r => r.days(r.c.revenue, r.avg(x => x.receivables)), [45, 90], false],
  ['النشاط والكفاءة', 'معدل دوران المخزون', 'تكلفة المبيعات ÷ متوسط المخزون', 'مرة', r => r.div(r.c.cogs, r.avg(x => x.inventory)), [6, 3], true],
  ['النشاط والكفاءة', 'متوسط فترة التخزين', '365 ÷ معدل دوران المخزون', 'يوم', r => r.days(r.c.cogs, r.avg(x => x.inventory)), [60, 120], false],
  ['النشاط والكفاءة', 'متوسط فترة السداد للموردين', 'متوسط الدائنين ÷ تكلفة المبيعات × 365', 'يوم', r => r.days(r.c.cogs, r.avg(x => x.payables)), [90, 120], false],
  ['النشاط والكفاءة', 'دورة التحويل النقدي', 'فترة التحصيل + فترة التخزين − فترة السداد', 'يوم', r => r.ccc(), [60, 120], false],

  ['الرفع المالي والملاءة', 'نسبة المديونية', 'إجمالي الالتزامات ÷ إجمالي الأصول', '%', r => r.pct(r.c.curL + r.c.nonCurL, r.c.curA + r.c.nonCurA), [50, 70], false],
  ['الرفع المالي والملاءة', 'الديون إلى حقوق الملكية', 'إجمالي الالتزامات ÷ حقوق الملكية', 'مرة', r => r.div(r.c.curL + r.c.nonCurL, r.c.totE), [1, 2], false],
  ['الرفع المالي والملاءة', 'نسبة الملكية', 'حقوق الملكية ÷ إجمالي الأصول', '%', r => r.pct(r.c.totE, r.c.curA + r.c.nonCurA), [50, 30], true],
  ['الرفع المالي والملاءة', 'تغطية الفوائد', 'الربح قبل الفوائد ÷ المصروفات المالية', 'مرة', r => r.div(r.c.netProfit + r.c.finance, r.c.finance), [5, 2], true],

  ['التدفقات النقدية', 'نسبة التدفق النقدي التشغيلي', 'صافي النقد التشغيلي ÷ الالتزامات المتداولة', 'مرة', r => r.ocf === null ? null : r.div(r.ocf, r.c.curL), [1, 0.5], true],
  ['التدفقات النقدية', 'جودة الأرباح', 'صافي النقد التشغيلي ÷ صافي الربح', 'مرة', r => (r.ocf === null || r.c.netProfit <= 0) ? null : r.div(r.ocf, r.c.netProfit), [1, 0.8], true],
  ['التدفقات النقدية', 'هامش التدفق النقدي', 'صافي النقد التشغيلي ÷ المبيعات', '%', r => r.ocf === null ? null : r.pct(r.ocf, r.c.revenue), [10, 5], true]
];

/* مؤشرات سنة واحدة. c = نموذج السنة، o = نموذج أول المدة (أو null)، ocf = صافي النقد التشغيلي (أو null) */
function ratioValues_(c, o, ocf) {
  const div = (a, b) => (b && Math.abs(b) > 0.0001) ? a / b : null;
  const r = {
    c, ocf,
    div: div,
    pct: (a, b) => { const v = div(a, b); return v === null ? null : v * 100; },
    avg: f => o ? (f(c) + f(o)) / 2 : f(c),
    days: (num, avgBal) => { const t = div(num, avgBal); return (t === null || t === 0) ? null : 365 / t; }
  };
  r.ccc = () => {
    const dso = r.days(c.revenue, r.avg(x => x.receivables));
    const dio = r.days(c.cogs, r.avg(x => x.inventory));
    const dpo = r.days(c.cogs, r.avg(x => x.payables));
    if (dso === null && dio === null) return null;
    return (dso || 0) + (dio || 0) - (dpo || 0);
  };
  return RATIO_DEFS_.map(d => { const v = d[4](r); return (v === null || !isFinite(v)) ? null : v; });
}

function rateRatio_(v, limits, higherBetter) {
  if (v === null) return 'غير متاح';
  if (higherBetter) return v >= limits[0] ? 'جيد' : (v >= limits[1] ? 'مقبول' : 'ضعيف');
  return v <= limits[0] ? 'جيد' : (v <= limits[1] ? 'مقبول' : 'ضعيف');
}

function benchmarkText_(d) {
  const u = d[3] === '%' ? '%' : (d[3] === 'يوم' ? ' يوم' : (d[3] === 'مرة' ? ' مرة' : ''));
  if (d[3] === 'مبلغ') return 'موجب';
  // بالكلمات بدل ≥ و≤ لأن المتصفح يعكس الرمزين في النص العربي
  return d[5][0] + u + (d[6] ? ' فأكثر' : ' فأقل');
}

function buildRatios_(m, mOpen, ocf, p, pOpen, ocfPrev) {
  const cur = ratioValues_(m, mOpen, ocf);
  const prev = p ? ratioValues_(p, pOpen, ocfPrev) : null;
  return RATIO_DEFS_.map((d, i) => ({
    cat: d[0], name: d[1], formula: d[2], unit: d[3],
    value: cur[i], prev: prev ? prev[i] : null,
    benchmark: benchmarkText_(d), eval: rateRatio_(cur[i], d[5], d[6]), higherBetter: d[6]
  }));
}

/* يضيف رقم السنة السابقة لكل سطر (بالاسم داخل نفس القسم). prevSecs = null يعني لا توجد مقارنة */
function mergeComparative_(curSecs, prevSecs) {
  return (curSecs || []).map(sec => {
    const ps = prevSecs ? prevSecs.find(x => x.section === sec.section) : null;
    const out = Object.assign({}, sec);
    if (sec.items) {
      out.items = sec.items.map(it => {
        if (it.heading) return it;
        const match = ps ? (ps.items || []).find(x => !x.heading && x.name === it.name) : null;
        return Object.assign({}, it, { prev: ps ? (match ? match.value : 0) : null });
      });
      // سطر موجود في السنة السابقة فقط (مثل استبعاد أصول)
      if (ps) (ps.items || []).forEach(x => {
        if (!x.heading && !out.items.some(i => i.name === x.name)) out.items.push({ name: x.name, value: 0, prev: x.value });
      });
    }
    if (sec.total !== undefined) out.prevTotal = ps ? ps.total : null;
    return out;
  });
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
  // null (لا توجد مقارنة) يُحفظ خانة فارغة
  const pc = v => (v === null || v === undefined) ? '' : v;
  const bsRows = [];
  bsObj.sections.forEach(sec => {
    (sec.items || []).forEach(it => bsRows.push([clientId, year, sec.section, it.name, it.value, 'بند', it.note || '', pc(it.prev)]));
    if (sec.total !== undefined) bsRows.push([clientId, year, sec.section, sec.totalName || 'الإجمالي', sec.total, 'إجمالي', '', pc(sec.prevTotal)]);
  });
  (bsObj.notes || []).forEach(n => bsRows.push([clientId, year, 'ملاحظات', n, '', 'ملاحظة']));
  const bsSh = getSheetByNameSafe_(ss, SHEET_NAMES.BS);
  [[6, 'النوع'], [7, 'إيضاح'], [8, 'السنة السابقة']].forEach(h => ensureHeaderCell_(bsSh, h[0], h[1]));
  replaceRows_(bsSh, match, bsRows);

  const isRows = isObj.items.map(it => [clientId, year, it.name, it.heading ? '' : it.value,
    it.heading ? 'عنوان' : (it.isFinal ? 'نهائي' : (it.isTotal ? 'إجمالي' : 'بند')), it.note || '', it.heading ? '' : pc(it.prev)]);
  const isSh = getSheetByNameSafe_(ss, SHEET_NAMES.IS);
  [[5, 'النوع'], [6, 'إيضاح'], [7, 'السنة السابقة']].forEach(h => ensureHeaderCell_(isSh, h[0], h[1]));
  replaceRows_(isSh, match, isRows);

  const cfRows = [];
  (cfObj.items || []).concat(cfObj.direct || []).forEach(sec => {
    (sec.items || []).forEach(it => cfRows.push([clientId, year, sec.section, it.name, it.heading ? '' : it.value, it.heading ? 'عنوان' : 'بند', it.heading ? '' : pc(it.prev)]));
    if (sec.total !== undefined) cfRows.push([clientId, year, sec.section, sec.totalName || ('إجمالي ' + sec.section), sec.total, 'إجمالي', pc(sec.prevTotal)]);
  });
  (cfObj.notes || []).forEach(n => cfRows.push([clientId, year, 'ملاحظات', n, '', 'ملاحظة']));
  const cfSh = getSheetByNameSafe_(ss, SHEET_NAMES.CF);
  [[6, 'النوع'], [7, 'السنة السابقة']].forEach(h => ensureHeaderCell_(cfSh, h[0], h[1]));
  replaceRows_(cfSh, match, cfRows);

  const ratioRows = ratios.map(r => [clientId, year, r.cat + ' - ' + r.name, pc(r.value), r.eval,
    r.unit || '', pc(r.prev), r.formula || '', r.benchmark || '', r.higherBetter === false ? 'أقل أفضل' : 'أعلى أفضل']);
  const ratSh = getSheetByNameSafe_(ss, SHEET_NAMES.RATIOS);
  [[6, 'الوحدة'], [7, 'السنة السابقة'], [8, 'المعادلة'], [9, 'المعيار'], [10, 'الاتجاه']].forEach(h => ensureHeaderCell_(ratSh, h[0], h[1]));
  replaceRows_(ratSh, match, ratioRows);
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

  // الصفوف القديمة (بدون عمود النوع أو السنة السابقة) تُميَّز بالاسم
  const pr = x => (x === '' || x === null || x === undefined) ? null : Number(x);
  const bsSections = [], bsNotes = [];
  bsRows.forEach(r => {
    const secName = String(r[2]), itemName = String(r[3]), val = r[4], kind = String(r[5] || '');
    if (kind === 'ملاحظة') { bsNotes.push(itemName); return; }
    let sec = bsSections.find(s => s.section === secName);
    if (!sec) { sec = { section: secName }; bsSections.push(sec); }
    const isTot = kind ? kind === 'إجمالي'
                       : (itemName === secName || itemName.includes('إجمالي') || secName.startsWith('إجمالي'));
    if (isTot) {
      sec.total = val; sec.totalName = itemName; sec.prevTotal = pr(r[7]);
      if (secName.startsWith('إجمالي')) sec.isTotal = true;
    } else {
      if (!sec.items) sec.items = [];
      sec.items.push({ name: itemName, value: val, note: r[6] || '', prev: pr(r[7]) });
    }
  });

  const isItems = isRows.map(r => {
    const name = String(r[2]), kind = String(r[4] || '');
    if (kind === 'عنوان') return { name: name, heading: true };
    if (kind) return { name: name, value: r[3], isTotal: kind === 'إجمالي' || kind === 'نهائي', isFinal: kind === 'نهائي', note: r[5] || '', prev: pr(r[6]) };
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
      sec.total = val; sec.totalName = itemName; sec.prevTotal = pr(r[6]);
      if (secName === 'صافي التغير في النقدية') sec.isTotal = true;
    } else {
      if (!sec.items) sec.items = [];
      sec.items.push(kind === 'عنوان' ? { name: itemName, heading: true } : { name: itemName, value: val, prev: pr(r[6]) });
    }
  });

  const ratios = ratRows.map(r => {
    const full = String(r[2]), i = full.indexOf(' - ');
    return { cat: i >= 0 ? full.slice(0, i) : full, name: i >= 0 ? full.slice(i + 3) : full,
             value: pr(r[3]), eval: r[4], unit: r[5] || '', prev: pr(r[6]), formula: r[7] || '',
             benchmark: r[8] || '', higherBetter: r[9] !== 'أقل أفضل' };
  });

  // قوائم محفوظة بإصدار سابق (بدون أرقام المقارنة) بينما يوجد ميزان للسنة السابقة
  const prevYear = String(Number(year) - 1);
  const noCmp = bsRows.every(r => r[7] === '' || r[7] === undefined) && ratRows.every(r => r[6] === '' || r[6] === undefined);
  const needsRefresh = (noCmp && getTrialBalance_(clientId, prevYear).length > 0) || ratRows.every(r => !r[5]);

  return { bs: { sections: bsSections, notes: bsNotes }, is: { items: isItems },
           cf: { items: cfSections, direct: cfDirect, notes: cfNotes }, ratios: ratios,
           years: { cur: String(year), prev: prevYear }, needsRefresh: needsRefresh };
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
