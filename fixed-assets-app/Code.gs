/**
 * نظام إدارة الأصول الثابتة — Fixed Assets Management
 * Google Sheets (قاعدة البيانات) + Google Apps Script (الخادم) + HTML (الواجهة)
 *
 * أسماء الأعمدة في الشيت تطابق أسماء حقول كيانات البيانات (Data entities) في
 * Microsoft Dynamics 365 Finance & Operations قدر الإمكان، لتسهيل الاستيراد والتصدير:
 *   - FixedAssets      ≈ Fixed assets V2      (FixedAssetV2Entity)
 *   - AssetBooks       ≈ Fixed asset books V2 (AssetBookV2Entity)
 *   - FAJournalLines   ≈ Fixed asset journal  (FixedAssetJournalLineEntity)
 *   - Vendors          ≈ Vendors V2           (VendVendorV2Entity)
 */

const APP_VERSION = '14.0.0';
const SETUP_VERSION = '14';
const SS_PROP = 'FA_SPREADSHEET_ID';
const HASH_ROUNDS = 300;
const AUTH_ERR = 'AUTH: ';
const AUDIT = ['CREATEDBY', 'CREATEDON', 'MODIFIEDBY', 'MODIFIEDON'];
const PERM_PAGES = ['assets', 'books', 'journal', 'transfers', 'gl', 'fiscal', 'budget', 'customers', 'banks', 'cashvouchers', 'invoices', 'items', 'depreciation', 'reports', 'masterdata', 'vendors', 'users', 'settings'];
const PERM_ACTIONS = ['view', 'add', 'edit', 'delete', 'approve', 'post'];
const USER_COLS = ['USERNAME', 'FULLNAME', 'EMAIL', 'PASSWORDHASH', 'SALT', 'ROLE', 'PERMISSIONS', 'STATUS',
  'MUSTCHANGE', 'LASTLOGIN', 'CREATEDBY', 'CREATEDON', 'MODIFIEDBY', 'MODIFIEDON'];
const SESSION_COLS = ['SESSIONID', 'USERNAME', 'CREATEDON', 'LASTSEEN', 'ACTIVE', 'CLIENT'];
const SETTING_DEFAULTS = {
  ORGNAME: 'الإدارة المالية', ORGEMAIL: '', CURRENCY: 'DJF', SESSIONTIMEOUT: '60',
  MAXSESSIONS: '3', DATEFORMAT: 'MDY', LOGO: '', JOURNALPREFIX: 'FAJ-', JOURNALSEQ: '0', TRANSFERPREFIX: 'TRF-', TRANSFERSEQ: '0', GLPREFIX: 'JV-', GLSEQ: '0', TRFJOURNALNAME: 'FAJ', RETAINEDEARNINGSACCOUNT: '', CLOSINGJOURNALNAME: 'YE', OPENINGBALANCEACCOUNT: '', BUDGETPREFIX: 'BG-', BUDGETSEQ: '0', RVPREFIX: 'RV-', RVSEQ: '0', PVPREFIX: 'PV-', PVSEQ: '0', SIPREFIX: 'SI-', SISEQ: '0', PIPREFIX: 'PI-', PISEQ: '0', SCNPREFIX: 'SCN-', SCNSEQ: '0', SDNPREFIX: 'SDN-', SDNSEQ: '0',
  PDNPREFIX: 'PDN-', PDNSEQ: '0', PCNPREFIX: 'PCN-', PCNSEQ: '0',
  SALESACCOUNT: '', PURCHASEACCOUNT: '', SALESTAXACCOUNT: '', PURCHTAXACCOUNT: '', DEFAULTTAXPERCENT: '0', SALESRETURNACCOUNT: '', PURCHRETURNACCOUNT: '',
  DEPJOURNALNAME: 'FADEP', DEPOFFSETACCOUNT: '', DEPMONTHS: '3'
};
const PUBLIC_SETTINGS = ['ORGNAME', 'ORGEMAIL', 'CURRENCY', 'SESSIONTIMEOUT', 'MAXSESSIONS', 'DATEFORMAT', 'LOGO', 'JOURNALPREFIX',
  'DEPJOURNALNAME', 'DEPOFFSETACCOUNT', 'DEPMONTHS', 'TRANSFERPREFIX', 'GLPREFIX', 'TRFJOURNALNAME', 'RETAINEDEARNINGSACCOUNT', 'CLOSINGJOURNALNAME', 'OPENINGBALANCEACCOUNT',
  'SALESACCOUNT', 'PURCHASEACCOUNT', 'SALESTAXACCOUNT', 'PURCHTAXACCOUNT', 'DEFAULTTAXPERCENT', 'SALESRETURNACCOUNT', 'PURCHRETURNACCOUNT'];

/* ============================== Schemas ============================== */

function o(v, ar, en) { return { v: v, ar: ar, en: en || v }; }
function F(k, ar, en, type, x) { return Object.assign({ k: k, ar: ar, en: en, type: type || 'text' }, x || {}); }

const ROLES = {
  SuperAdmin: { ar: 'مدير عام', en: 'Super admin' },
  Admin: { ar: 'مدير', en: 'Admin' },
  Accountant: { ar: 'محاسب', en: 'Accountant' },
  Viewer: { ar: 'مستخدم عرض', en: 'Viewer' }
};

const ASSET_TYPES = [o('Tangible', 'ملموس'), o('Intangible', 'غير ملموس'), o('Financial', 'مالي'),
  o('Land', 'أرض'), o('Goodwill', 'شهرة'), o('Other', 'أخرى')];

const BOOK_STATUS = [o('NoAcquisition', 'لم يتم الاقتناء', 'Not yet acquired'), o('Open', 'مفتوح'),
  o('Suspended', 'معلّق'), o('Closed', 'مغلق'), o('Sold', 'مباع'), o('Scrapped', 'مُخرّد', 'Scrapped')];

const TRANS_TYPES = [
  o('Acquisition', 'اقتناء', 'Acquisition'), o('AcquisitionAdj', 'تسوية اقتناء', 'Acquisition adjustment'),
  o('Depreciation', 'إهلاك', 'Depreciation'), o('DepreciationAdj', 'تسوية إهلاك', 'Depreciation adjustment'),
  o('ExtraordinaryDepreciation', 'إهلاك استثنائي', 'Extraordinary depreciation'),
  o('Revaluation', 'إعادة تقييم', 'Revaluation'), o('WriteUpAdj', 'تسوية زيادة القيمة', 'Write-up adjustment'),
  o('WriteDownAdj', 'تسوية تخفيض القيمة', 'Write-down adjustment'),
  o('DisposalSale', 'استبعاد - بيع', 'Disposal - sale'), o('DisposalScrap', 'استبعاد - تخريد', 'Disposal - scrap')
];

const JOURNAL_STATUS = [o('Draft', 'مسودة'), o('Approved', 'معتمد'), o('Posted', 'مُرحّل')];
const GL_SOURCES = [o('Manual', 'يدوي', 'Manual'), o('FAJournal', 'يومية الأصول', 'Fixed asset journal'), o('Transfer', 'نقل أصول', 'Asset transfer'),
  o('YearEnd', 'إقفال السنة', 'Year-end closing'), o('CashVoucher', 'سند قبض / صرف', 'Receipt / payment'), o('Invoice', 'فاتورة', 'Invoice')];
const INVOICE_TYPES = [o('Sales', 'فاتورة مبيعات', 'Sales invoice'), o('Purchase', 'فاتورة مشتريات', 'Purchase invoice'),
  o('SalesCredit', 'إشعار دائن للعميل (مرتجع مبيعات / خصم)', 'Customer credit note'), o('SalesDebit', 'إشعار مدين للعميل', 'Customer debit note'),
  o('PurchaseDebit', 'إشعار مدين للمورد (مرتجع مشتريات)', 'Vendor debit note'), o('PurchaseCredit', 'إشعار دائن من المورد', 'Vendor credit note')];
/** خصائص كل نوع: الجانب (مبيعات/مشتريات)، هل يخفّض رصيد الطرف، الترقيم، اليومية، الاسم. */
const INV_KIND = {
  Sales: { side: 'Sales', credit: 0, seq: ['SISEQ', 'SIPREFIX', 'SI-'], jn: 'SI', ar: 'فاتورة مبيعات' },
  Purchase: { side: 'Purchase', credit: 0, seq: ['PISEQ', 'PIPREFIX', 'PI-'], jn: 'PI', ar: 'فاتورة مشتريات' },
  SalesCredit: { side: 'Sales', credit: 1, seq: ['SCNSEQ', 'SCNPREFIX', 'SCN-'], jn: 'CN', ar: 'إشعار دائن للعميل' },
  SalesDebit: { side: 'Sales', credit: 0, seq: ['SDNSEQ', 'SDNPREFIX', 'SDN-'], jn: 'DN', ar: 'إشعار مدين للعميل' },
  PurchaseDebit: { side: 'Purchase', credit: 1, seq: ['PDNSEQ', 'PDNPREFIX', 'PDN-'], jn: 'DN', ar: 'إشعار مدين للمورد' },
  PurchaseCredit: { side: 'Purchase', credit: 0, seq: ['PCNSEQ', 'PCNPREFIX', 'PCN-'], jn: 'CN', ar: 'إشعار دائن من المورد' }
};
const BUDGET_TYPES = [o('Expense', 'مصروفات', 'Expense'), o('Revenue', 'إيرادات', 'Revenue')];
const BUDGET_STATUS = [o('Draft', 'مسودة', 'Draft'), o('Approved', 'معتمدة', 'Approved'), o('Posted', 'مُرحّلة', 'Posted')];
const CASH_TYPES = [o('Receipt', 'سند قبض', 'Receipt'), o('Payment', 'سند صرف', 'Payment')];
const CASH_PARTY = [o('Customer', 'عميل', 'Customer'), o('Vendor', 'مورد', 'Vendor'), o('Ledger', 'حساب أستاذ', 'Ledger account')];
const CASH_STATUS = [o('Draft', 'مسودة', 'Draft'), o('Posted', 'مُرحّل', 'Posted'), o('Cancelled', 'ملغى', 'Cancelled')];
const PAY_METHODS = [o('Cash', 'نقداً', 'Cash'), o('Check', 'شيك', 'Cheque'), o('Transfer', 'تحويل بنكي', 'Bank transfer')];
const BUDGET_PERIODS = [o('Annual', 'سنوي (يوزّع على 12 شهراً)', 'Annual (spread over 12 months)'), o('Monthly', 'شهري (لكل شهر)', 'Monthly (each month)')];
const ACCOUNT_KINDS = [o('Ledger', 'حساب أستاذ', 'Ledger'), o('Customer', 'عميل', 'Customer'), o('Vendor', 'مورد', 'Vendor'), o('Bank', 'صندوق / بنك', 'Bank')];
const FISCAL_STATUS = [o('Open', 'مفتوحة', 'Open'), o('Closed', 'مقفلة', 'Closed')];

const ACCOUNT_TYPES = [o('ProfitAndLoss', 'أرباح وخسائر', 'Profit and loss'), o('Revenue', 'إيرادات', 'Revenue'),
  o('Expense', 'مصروفات', 'Expense'), o('BalanceSheet', 'الميزانية العمومية', 'Balance sheet'), o('Asset', 'أصول', 'Asset'),
  o('Liability', 'خصوم', 'Liability'), o('Equity', 'حقوق ملكية', 'Equity'), o('Total', 'إجمالي', 'Total'),
  o('Reporting', 'تقارير', 'Reporting')];
const CASHFLOW_CATS = [o('Cash', 'نقدية وما في حكمها', 'Cash and cash equivalents'), o('Operating', 'أنشطة تشغيلية', 'Operating activities'),
  o('Investing', 'أنشطة استثمارية', 'Investing activities'), o('Financing', 'أنشطة تمويلية', 'Financing activities')];
const ACCOUNT_LEVELS = [o('Header', 'رئيسي', 'Header'), o('Posting', 'فرعي', 'Posting')];

/**
 * القوائم الأساسية — كل نوع في صفحة (شيت) مستقلة: كود + اسم + كود الأب.
 * parent: النوع الأعلى في الشجرة.
 */
const LK = [
  { v: 'OperatingUnit', sheet: 'OperatingUnits', ar: 'الوحدات الرئيسية', en: 'Operating units', d365: 'A_OperatingUnit' },
  { v: 'Sub1_OperatingUnit', sheet: 'Sub1OperatingUnits', parent: 'OperatingUnit', req: 1, ar: 'الوحدات الفرعية 1', en: 'Sub 1 operating units', d365: 'A_Sub1_OperatingUnit' },
  { v: 'Sub2_OperatingUnit', sheet: 'Sub2OperatingUnits', parent: 'Sub1_OperatingUnit', req: 1, ar: 'الوحدات الفرعية 2', en: 'Sub 2 operating units', d365: 'A_Sub2_OperatingUnit' },
  { v: 'Location', sheet: 'Locations', ar: 'المواقع', en: 'Locations' },
  { v: 'Building', sheet: 'Buildings', parent: 'Location', ar: 'المباني', en: 'Buildings' },
  { v: 'Room', sheet: 'Rooms', parent: 'Building', ar: 'الغرف', en: 'Rooms' },
  { v: 'Department', sheet: 'Departments', ar: 'الإدارات / الأقسام', en: 'Departments' },
  { v: 'Employee', sheet: 'Employees', ar: 'الموظفون المسؤولون', en: 'Responsible employees' },
  { v: 'JournalName', sheet: 'JournalNames', ar: 'أسماء اليوميات', en: 'Journal names' },
  { v: 'BankGroup', sheet: 'BankGroups', ar: 'مجموعات الصناديق والبنوك', en: 'Cash & bank groups' },
  { v: 'ItemUnit', sheet: 'ItemUnits', ar: 'وحدات الأصناف', en: 'Units of measure', d365: 'UnitSymbol' }
];
const LKMAP = {};
LK.forEach(function (t) { LKMAP[t.v] = t; });

/** سلاسل الشجرة (من الأدنى إلى الأعلى): المستخدم يختار الأدنى والنظام يملأ الأعلى. */
const UNIT_CHAIN = [['SUB2_OPERATINGUNIT', 'Sub2_OperatingUnit'], ['SUB1_OPERATINGUNIT', 'Sub1_OperatingUnit'], ['OPERATINGUNIT', 'OperatingUnit']];
const OFFSET_CHAIN = [['OFFSET_SUB2_OPERATINGUNIT', 'Sub2_OperatingUnit'], ['OFFSET_SUB1_OPERATINGUNIT', 'Sub1_OperatingUnit'], ['OFFSET_OPERATINGUNIT', 'OperatingUnit']];
const LOCATION_CHAIN = [['ROOMNUMBER', 'Room'], ['BUILDING', 'Building'], ['LOCATION', 'Location']];
const TO_UNIT_CHAIN = UNIT_CHAIN.map(function (c) { return ['TO_' + c[0], c[1]]; });
const TO_LOCATION_CHAIN = LOCATION_CHAIN.map(function (c) { return ['TO_' + c[0], c[1]]; });
/** الحقول التي يغيّرها النقل على الأصل. */
const TRF_UNIT = ['OPERATINGUNIT', 'SUB1_OPERATINGUNIT', 'SUB2_OPERATINGUNIT'];
const TRF_LOC = ['LOCATION', 'BUILDING', 'ROOMNUMBER'];
const TRF_FIELDS = TRF_UNIT.concat(TRF_LOC).concat(['DEPARTMENT', 'RESPONSIBLE']);

/** حقول الموقع ← المبنى ← الغرفة (الغرفة تملأ المبنى والمبنى يملأ الموقع). */
function locFields_(prefix, sec, grp, calc, listKeys) {
  listKeys = listKeys || {};
  return [
    F(prefix + 'LOCATION', 'الموقع', 'Location', 'lookup', { lk: 'Location', calc: calc ? 1 : 0, sec: sec, grp: grp, list: listKeys.LOCATION ? 1 : 0 }),
    F(prefix + 'BUILDING', 'المبنى', 'Building', 'lookup', { lk: 'Building', up: prefix + 'LOCATION', calc: calc ? 1 : 0, sec: sec, grp: grp }),
    F(prefix + 'ROOMNUMBER', 'الغرفة', 'Room', 'lookup', { lk: 'Room', up: prefix + 'BUILDING', calc: calc ? 1 : 0, sec: sec, grp: grp, list: listKeys.ROOMNUMBER ? 1 : 0 })
  ];
}

/** حقول الوحدات الثلاثة: المستخدم يسجّل الفرعية 2 والباقي يُكتب تلقائياً. */
function unitFields_(prefix, sec, grp, x) {
  x = x || {};
  return [
    F(prefix + 'OPERATINGUNIT', 'الوحدة الرئيسية', 'A_A_OperatingUnit', 'lookup', Object.assign({ lk: 'OperatingUnit', calc: 1, sec: sec, grp: grp,
      aliases: prefix ? [] : ['A_OperatingUnit', 'A_A_OperatingUnit', 'OperatingUnit'] }, x.ou || {})),
    F(prefix + 'SUB1_OPERATINGUNIT', 'الوحدة الفرعية 1', 'A_Sub1_OperatingUnit', 'lookup', { lk: 'Sub1_OperatingUnit', up: prefix + 'OPERATINGUNIT', calc: 1, sec: sec, grp: grp,
      aliases: prefix ? [] : ['A_Sub1_OperatingUnit', 'Sub1_OperatingUnit'] }),
    F(prefix + 'SUB2_OPERATINGUNIT', 'الوحدة الفرعية 2', 'A_Sub2_OperatingUnit', 'lookup', { lk: 'Sub2_OperatingUnit', up: prefix + 'SUB1_OPERATINGUNIT',
      calc: x.calc ? 1 : 0, sec: sec, grp: grp, aliases: prefix ? [] : ['A_Sub2_OperatingUnit', 'Sub2_OperatingUnit'] })
  ];
}

const SCHEMAS = {
  assets: {
    sheet: 'FixedAssets', perm: 'assets', key: ['FIXEDASSETNUMBER'], title: { ar: 'الأصول الثابتة', en: 'Fixed assets' },
    sections: [
      { id: 'general', ar: 'عام', en: 'General', groups: [
        { id: 'id', ar: 'التعريف', en: 'Identification' }, { id: 'desc', ar: 'الوصف', en: 'Description' },
        { id: 'type', ar: 'النوع', en: 'Type' }, { id: 'qty', ar: 'الكمية', en: 'Quantity' }] },
      { id: 'dims', ar: 'الوحدات التشغيلية (الأبعاد المالية)', en: 'Operating units (financial dimensions)', groups: [
        { id: 'dims', ar: 'سجّل الوحدة الفرعية 2 وتُكتب الوحدات الأعلى تلقائياً', en: 'Enter sub-unit 2; higher units fill automatically' }] },
      { id: 'tech', ar: 'معلومات فنية', en: 'Technical information', groups: [
        { id: 'model', ar: 'الموديل', en: 'Model' }] },
      { id: 'loc', ar: 'الموقع والمسؤولية', en: 'Location & responsibility', groups: [
        { id: 'loc', ar: 'الموقع (اختر الغرفة ويُكتب المبنى والموقع)', en: 'Location (pick the room; building and location fill in)' },
        { id: 'emp', ar: 'المسؤول', en: 'Responsible' }] }
    ],
    fields: [
      F('FIXEDASSETNUMBER', 'رقم الأصل', 'Number', 'text', { req: 1, key: 1, sec: 'general', grp: 'id', list: 1, sum: 1,
        aliases: ['Fixed asset number', 'Asset ID', 'ASSETID', 'Fixed asset'] }),
      F('FIXEDASSETGROUPID', 'مجموعة الأصل', 'Fixed asset group', 'lookup', { lk: 'AssetGroup', req: 1, sec: 'general', grp: 'id', list: 1,
        aliases: ['Fixed asset group', 'Asset group', 'ASSETGROUP', 'GROUPID'] }),
      F('ASSETACCOUNT', 'حساب الأصل (من المجموعة)', 'Asset account', 'lookup', { lk: '@allAccounts', calc: 1, sec: 'general', grp: 'id' }),
      F('NAME', 'اسم الأصل', 'Name', 'text', { req: 1, sec: 'general', grp: 'desc', list: 1, sum: 1, aliases: ['Asset name', 'Fixed asset name'] }),
      F('ASSETTYPE', 'النوع', 'Type', 'select', { opts: ASSET_TYPES, dflt: 'Tangible', sec: 'general', grp: 'type', list: 1, sum: 1,
        aliases: ['Asset type', 'TYPE'] }),
      F('QUANTITY', 'الكمية', 'Quantity', 'number', { dec: 2, dflt: 1, sec: 'general', grp: 'qty', list: 1 })
    ].concat(unitFields_('', 'dims', 'dims', { ou: { list: 1, sum: 1 } })).concat([
      F('MODEL', 'الموديل', 'Model', 'text', { sec: 'tech', grp: 'model', list: 1, sum: 1 }),
      F('GUARANTEEDATE', 'تاريخ الضمان', 'Guarantee date', 'date', { sec: 'tech', grp: 'model', sum: 1,
        aliases: ['Warranty date', 'WARRANTYDATE'] }),
      F('LOCATION', 'الموقع', 'Location', 'lookup', { lk: 'Location', sec: 'loc', grp: 'loc', list: 1, sum: 1,
        aliases: ['Location ID', 'LOCATIONID', 'Location memo', 'LOCATIONMEMO'] }),
      F('BUILDING', 'المبنى', 'Building', 'lookup', { lk: 'Building', up: 'LOCATION', sec: 'loc', grp: 'loc', list: 1, aliases: ['Building ID'] }),
      F('ROOMNUMBER', 'رقم الغرفة', 'Room number', 'lookup', { lk: 'Room', up: 'BUILDING', sec: 'loc', grp: 'loc', aliases: ['Room'] }),
      F('DEPARTMENT', 'الإدارة / القسم', 'Department', 'lookup', { lk: 'Department', sec: 'loc', grp: 'emp', list: 1 }),
      F('RESPONSIBLE', 'الموظف المسؤول', 'Employee', 'lookup', { lk: 'Employee', sec: 'loc', grp: 'emp', list: 1, sum: 1,
        aliases: ['Worker responsible', 'WORKERRESPONSIBLE', 'Responsible', 'Employee', 'WORKERRESPONSIBLEPERSONNELNUMBER'] })
    ])
  },

  books: {
    sheet: 'AssetBooks', perm: 'books', key: ['FIXEDASSETNUMBER', 'BOOKID'], title: { ar: 'دفاتر الأصول', en: 'Books' },
    sections: [
      { id: 'general', ar: 'عام', en: 'General', groups: [
        { id: 'book', ar: 'الدفتر', en: 'Book' }, { id: 'acq', ar: 'الاقتناء', en: 'Acquisition' },
        { id: 'bal', ar: 'الرصيد', en: 'Balance' }] },
      { id: 'dep', ar: 'الإهلاك', en: 'Depreciation', groups: [
        { id: 'life', ar: 'العمر الإنتاجي', en: 'Service life' }, { id: 'run', ar: 'الإهلاك', en: 'Depreciation' }] },
      { id: 'dims', ar: 'الأبعاد المالية الافتراضية (من الأصل)', en: 'Default financial dimensions (from the asset)', groups: [
        { id: 'dims', ar: 'الأبعاد المالية', en: 'Financial dimensions' }] }
    ],
    fields: [
      F('FIXEDASSETNUMBER', 'رقم الأصل', 'Fixed asset number', 'lookup', { lk: '@assets', req: 1, key: 1, sec: 'general', grp: 'book', list: 1,
        aliases: ['Fixed asset', 'Asset ID', 'ASSETID', 'Number'] }),
      F('BOOKID', 'رقم الدفتر', 'Book', 'text', { req: 1, key: 1, dflt: 'General', sec: 'general', grp: 'book', list: 1, sum: 1,
        aliases: ['Book ID', 'Book number', 'BOOKNUMBER'] }),
      F('ACQUISITIONDATE', 'تاريخ الاقتناء', 'Acquisition date', 'date', { sec: 'general', grp: 'acq', list: 1, sum: 1 }),
      F('ACQUISITIONPRICE', 'سعر الاقتناء', 'Acquisition price', 'number', { dec: 3, sec: 'general', grp: 'acq', list: 1 }),
      F('STATUS', 'الحالة', 'Status', 'select', { opts: BOOK_STATUS, dflt: 'NoAcquisition', sec: 'general', grp: 'bal', list: 1, sum: 1,
        aliases: ['Book status', 'ASSETSTATUS'] }),
      F('SERVICELIFE', 'العمر الإنتاجي (سنوات)', 'Service life', 'number', { dec: 2, sec: 'dep', grp: 'life', list: 1 }),
      F('DEPRECIATIONPERIODS', 'فترات الإهلاك (شهور)', 'Depreciation periods', 'number', { dec: 0, sec: 'dep', grp: 'life', list: 1,
        aliases: ['LIFETIME', 'Lifetime'] }),
      F('DEPRECIATIONPERIODSREMAINING', 'فترات الإهلاك المتبقية', 'Depreciation periods remaining', 'number', { dec: 2, sec: 'dep', grp: 'life', list: 1,
        aliases: ['LIFETIMEREST', 'Lifetime remaining', 'REMAININGLIFE'] }),
      F('SCRAPVALUE', 'القيمة المتوقعة للخردة', 'Expected scrap value', 'number', { dec: 3, sec: 'dep', grp: 'life',
        aliases: ['Scrap value', 'EXPECTEDSCRAPVALUE'] }),
      F('ACCUMULATEDDEPRECIATION', 'مجمع الإهلاك', 'Accumulated depreciation', 'number', { dec: 3, sec: 'general', grp: 'bal', list: 1,
        aliases: ['Accumulated depreciation', 'ACCUMDEPRECIATION', 'DEPRECIATIONAMOUNT'] }),
      F('NETBOOKVALUE', 'صافي القيمة الدفترية', 'Net book value', 'number', { dec: 3, calc: 1, sec: 'general', grp: 'bal', list: 1, sum: 1,
        aliases: ['Net book value', 'NBV'] }),
      F('DEPRECIATIONSTARTDATE', 'تاريخ تشغيل الإهلاك', 'Depreciation run date', 'date', { sec: 'dep', grp: 'run',
        aliases: ['DEPRECIATIONRUNDATE', 'Depreciation start date'] }),
      F('LASTDEPRECIATIONDATE', 'تاريخ آخر إهلاك', 'Date when depreciation was last run', 'date', { sec: 'dep', grp: 'run', list: 1,
        aliases: ['Last depreciation date', 'LASTDEPRECIATIONRUNDATE'] })
    ].concat(unitFields_('', 'dims', 'dims', { calc: 1, ou: { list: 1 } })).concat([
      F('DEFAULTDIMENSIONDISPLAYVALUE', 'قيمة الأبعاد الافتراضية', 'Default dimension display value', 'text', { calc: 1, sec: 'dims', grp: 'dims',
        aliases: ['Default dimension', 'DEFAULTDIMENSION', 'Financial dimensions'] }),
      F('OPENINGVOUCHER', 'قيد الرصيد الافتتاحي', 'Opening voucher', 'text', { sys: 1, sec: 'general', grp: 'bal' }),
      F('OPENINGCOST', 'التكلفة الافتتاحية المُقيّدة', 'Opening cost posted', 'number', { dec: 3, sys: 1, sec: 'general', grp: 'bal' }),
      F('OPENINGACCUM', 'مجمع الإهلاك الافتتاحي المُقيّد', 'Opening accumulated depreciation posted', 'number', { dec: 3, sys: 1, sec: 'general', grp: 'bal' })
    ])
  },

  journal: {
    sheet: 'FAJournalLines', perm: 'journal', key: ['LINEID'], autoId: 'LINEID', title: { ar: 'يومية الأصول الثابتة', en: 'Fixed asset journal' },
    sections: [
      { id: 'line', ar: 'السطر', en: 'Line', groups: [
        { id: 'batch', ar: 'اليومية', en: 'Journal' }, { id: 'asset', ar: 'الأصل', en: 'Fixed asset' },
        { id: 'amt', ar: 'المبلغ', en: 'Amount' }] },
      { id: 'offset', ar: 'الحساب المقابل (من دليل الحسابات)', en: 'Offset account (chart of accounts)', groups: [
        { id: 'offset', ar: 'الحساب المقابل', en: 'Offset account' }, { id: 'odims', ar: 'الوحدات (سجّل الفرعية 2)', en: 'Units (enter sub-unit 2)' }] },
      { id: 'wf', ar: 'سير العمل', en: 'Workflow', groups: [{ id: 'wf', ar: 'الحالة', en: 'Status' }] }
    ],
    fields: [
      F('LINEID', 'رقم السطر', 'Line ID', 'text', { key: 1, sys: 1, sec: 'wf', grp: 'wf' }),
      F('JOURNALBATCHNUMBER', 'رقم دفعة اليومية', 'Journal batch number', 'text', { sec: 'line', grp: 'batch', list: 1,
        aliases: ['Journal batch', 'JOURNALNUM', 'Journal'] }),
      F('TRANSDATE', 'التاريخ', 'Date', 'date', { req: 1, dflt: '@today', sec: 'line', grp: 'batch', list: 1,
        aliases: ['Transaction date', 'TRANSACTIONDATE', 'ACCOUNTINGDATE'] }),
      F('JOURNALNAME', 'اسم اليومية', 'Journal name', 'lookup', { lk: 'JournalName', req: 1, sec: 'line', grp: 'batch', list: 1 }),
      F('TRANSACTIONTYPE', 'نوع المعاملة', 'Transaction type', 'select', { opts: TRANS_TYPES, req: 1, sec: 'line', grp: 'asset', list: 1,
        aliases: ['TRANSTYPE', 'Fixed asset transaction type'] }),
      F('FIXEDASSETNUMBER', 'رقم الأصل', 'Fixed asset number', 'lookup', { lk: '@assets', req: 1, sec: 'line', grp: 'asset', list: 1,
        aliases: ['Fixed asset', 'Asset ID', 'ACCOUNTDISPLAYVALUE', 'ASSETID'] }),
      F('ASSETNAME', 'اسم الأصل', 'Asset name', 'text', { calc: 1, sec: 'line', grp: 'asset', list: 1, aliases: ['Name', 'Fixed asset name'] }),
      F('BOOKID', 'الدفتر', 'Book', 'lookup', { lk: '@books', dflt: 'General', sec: 'line', grp: 'asset', list: 1, aliases: ['Book ID', 'FIXEDASSETBOOKID'] }),
      F('LEDGERACCOUNT', 'حساب الأصل في القيد (من المجموعة)', 'Fixed asset ledger account', 'lookup', { lk: '@allAccounts', calc: 1, sec: 'line', grp: 'asset' }),
      F('TEXT', 'الوصف', 'Description', 'text', { sec: 'line', grp: 'amt', list: 1, aliases: ['Description', 'TRANSTEXT', 'Transaction text'] }),
      F('DEBITAMOUNT', 'مدين', 'Debit', 'number', { dec: 3, sec: 'line', grp: 'amt', list: 1, aliases: ['AMOUNTCURDEBIT', 'Debit amount'] }),
      F('CREDITAMOUNT', 'دائن', 'Credit', 'number', { dec: 3, sec: 'line', grp: 'amt', list: 1, aliases: ['AMOUNTCURCREDIT', 'Credit amount'] }),
      F('DEPRECIATIONMONTHS', 'عدد شهور الإهلاك', 'Depreciation months', 'number', { dec: 0, sec: 'line', grp: 'amt',
        aliases: ['Depreciation periods', 'PERIODS'] }),
      F('VENDORACCOUNT', 'المورد (يملأ حسابه تلقائياً)', 'Vendor', 'lookup', { lk: '@vendors', sec: 'offset', grp: 'offset', aliases: ['Vendor account'] }),
      F('OFFSETACCOUNT', 'رقم الحساب المقابل', 'Offset main account', 'lookup', { lk: '@accounts', sec: 'offset', grp: 'offset',
        aliases: ['Offset account number', 'OFFSETMAINACCOUNT', 'MAINACCOUNT'] })
    ].concat(unitFields_('OFFSET_', 'offset', 'odims')).concat([
      F('OFFSETACCOUNTDISPLAYVALUE', 'الحساب المقابل', 'Offset account', 'text', { calc: 1, sec: 'offset', grp: 'offset', list: 1,
        aliases: ['Offset account display value'] }),
      F('STATUS', 'الحالة', 'Status', 'select', { opts: JOURNAL_STATUS, sys: 1, sec: 'wf', grp: 'wf', list: 1 }),
      F('APPROVEDBY', 'اعتمد بواسطة', 'Approved by', 'text', { sys: 1, sec: 'wf', grp: 'wf' }),
      F('APPROVEDON', 'تاريخ الاعتماد', 'Approved on', 'text', { sys: 1, sec: 'wf', grp: 'wf' }),
      F('POSTEDBY', 'رحّل بواسطة', 'Posted by', 'text', { sys: 1, sec: 'wf', grp: 'wf' }),
      F('POSTEDON', 'تاريخ الترحيل', 'Posted on', 'text', { sys: 1, sec: 'wf', grp: 'wf' }),
      F('GLVOUCHER', 'رقم القيد المحاسبي', 'GL voucher', 'text', { sys: 1, sec: 'wf', grp: 'wf', list: 1 })
    ])
  },

  transfers: {
    sheet: 'AssetTransfers', perm: 'transfers', key: ['LINEID'], autoId: 'LINEID', title: { ar: 'نقل الأصول', en: 'Fixed asset transfers' },
    sections: [
      { id: 'hdr', ar: 'مستند النقل', en: 'Transfer document', groups: [
        { id: 'doc', ar: 'المستند', en: 'Document' }, { id: 'asset', ar: 'الأصل', en: 'Fixed asset' }] },
      { id: 'from', ar: 'من (الوضع الحالي)', en: 'From (current)', groups: [
        { id: 'fu', ar: 'الوحدات', en: 'Units' }, { id: 'fl', ar: 'الموقع', en: 'Location' }, { id: 'fr', ar: 'المسؤولية', en: 'Responsibility' }] },
      { id: 'to', ar: 'إلى (الوجهة الجديدة)', en: 'To (new destination)', groups: [
        { id: 'tu', ar: 'الوحدات (سجّل الفرعية 2)', en: 'Units (enter sub-unit 2)' }, { id: 'tl', ar: 'الموقع (اختر الغرفة)', en: 'Location (pick the room)' },
        { id: 'tr', ar: 'المسؤولية', en: 'Responsibility' }] },
      { id: 'wf', ar: 'سير العمل', en: 'Workflow', groups: [{ id: 'wf', ar: 'الحالة', en: 'Status' }] }
    ],
    fields: [
      F('LINEID', 'رقم السطر', 'Line ID', 'text', { key: 1, sys: 1, sec: 'wf', grp: 'wf' }),
      F('TRANSFERNUMBER', 'رقم مستند النقل', 'Transfer number', 'text', { sec: 'hdr', grp: 'doc', list: 1 }),
      F('TRANSFERDATE', 'تاريخ النقل', 'Transfer date', 'date', { req: 1, dflt: '@today', sec: 'hdr', grp: 'doc', list: 1 }),
      F('REASON', 'سبب النقل', 'Reason', 'text', { sec: 'hdr', grp: 'doc', list: 1, aliases: ['Description', 'Notes'] }),
      F('FIXEDASSETNUMBER', 'رقم الأصل', 'Fixed asset number', 'lookup', { lk: '@assets', req: 1, sec: 'hdr', grp: 'asset', list: 1,
        aliases: ['Fixed asset', 'Asset ID'] }),
      F('ASSETNAME', 'اسم الأصل', 'Asset name', 'text', { calc: 1, sec: 'hdr', grp: 'asset', list: 1 })
    ]
      .concat(unitFields_('FROM_', 'from', 'fu', { calc: 1 })).concat(locFields_('FROM_', 'from', 'fl', true, { LOCATION: 1, ROOMNUMBER: 1 }))
      .concat([
        F('FROM_DEPARTMENT', 'الإدارة / القسم', 'Department', 'lookup', { lk: 'Department', calc: 1, sec: 'from', grp: 'fr' }),
        F('FROM_RESPONSIBLE', 'الموظف المسؤول', 'Employee', 'lookup', { lk: 'Employee', calc: 1, sec: 'from', grp: 'fr', list: 1 })
      ])
      .concat(unitFields_('TO_', 'to', 'tu')).concat(locFields_('TO_', 'to', 'tl', false, { LOCATION: 1, ROOMNUMBER: 1 }))
      .concat([
        F('TO_DEPARTMENT', 'الإدارة / القسم', 'Department', 'lookup', { lk: 'Department', sec: 'to', grp: 'tr' }),
        F('TO_RESPONSIBLE', 'الموظف المسؤول', 'Employee', 'lookup', { lk: 'Employee', sec: 'to', grp: 'tr', list: 1 }),
        F('STATUS', 'الحالة', 'Status', 'select', { opts: JOURNAL_STATUS, sys: 1, sec: 'wf', grp: 'wf', list: 1 }),
        F('APPROVEDBY', 'اعتمد بواسطة', 'Approved by', 'text', { sys: 1, sec: 'wf', grp: 'wf' }),
        F('APPROVEDON', 'تاريخ الاعتماد', 'Approved on', 'text', { sys: 1, sec: 'wf', grp: 'wf' }),
        F('POSTEDBY', 'رحّل بواسطة', 'Posted by', 'text', { sys: 1, sec: 'wf', grp: 'wf' }),
        F('POSTEDON', 'تاريخ الترحيل', 'Posted on', 'text', { sys: 1, sec: 'wf', grp: 'wf' }),
      F('GLVOUCHER', 'رقم القيد المحاسبي', 'GL voucher', 'text', { sys: 1, sec: 'wf', grp: 'wf', list: 1 })
      ])
  },

  glheaders: {
    sheet: 'GLJournals', perm: 'gl', key: ['VOUCHER'], title: { ar: 'القيود المحاسبية', en: 'Journal entries' },
    sections: [{ id: 'hdr', ar: 'رأس القيد', en: 'Voucher header', groups: [
      { id: 'h1', ar: 'القيد', en: 'Voucher' }, { id: 'h2', ar: 'البيان', en: 'Description' }, { id: 'h3', ar: 'المصدر', en: 'Source' }] }],
    fields: [
      F('VOUCHER', 'رقم القيد', 'Voucher', 'text', { key: 1, sec: 'hdr', grp: 'h1', list: 1, aliases: ['Journal batch number', 'JOURNALBATCHNUMBER'] }),
      F('TRANSDATE', 'التاريخ', 'Date', 'date', { req: 1, dflt: '@today', sec: 'hdr', grp: 'h1', list: 1, aliases: ['Transaction date', 'ACCOUNTINGDATE'] }),
      F('JOURNALNAME', 'نوع اليومية', 'Journal name', 'lookup', { lk: 'JournalName', req: 1, sec: 'hdr', grp: 'h1', list: 1 }),
      F('DESCRIPTION', 'البيان', 'Description', 'text', { sec: 'hdr', grp: 'h2', list: 1, aliases: ['Text', 'TEXT'] }),
      F('REFERENCE', 'المرجع', 'Reference', 'text', { sec: 'hdr', grp: 'h2', list: 1, aliases: ['Document', 'DOCUMENT'] }),
      F('SOURCE', 'المصدر', 'Source', 'select', { opts: GL_SOURCES, sys: 1, sec: 'hdr', grp: 'h3', list: 1 }),
      F('SOURCEREF', 'مستند المصدر', 'Source document', 'text', { sys: 1, sec: 'hdr', grp: 'h3', list: 1 }),
      F('TOTALDEBIT', 'إجمالي المدين', 'Total debit', 'number', { dec: 3, sys: 1, sec: 'hdr', grp: 'h3', list: 1 }),
      F('TOTALCREDIT', 'إجمالي الدائن', 'Total credit', 'number', { dec: 3, sys: 1, sec: 'hdr', grp: 'h3', list: 1 }),
      F('STATUS', 'الحالة', 'Status', 'select', { opts: JOURNAL_STATUS, sys: 1, sec: 'hdr', grp: 'h3', list: 1 }),
      F('APPROVEDBY', 'اعتمد بواسطة', 'Approved by', 'text', { sys: 1, sec: 'hdr', grp: 'h3' }),
      F('POSTEDBY', 'رحّل بواسطة', 'Posted by', 'text', { sys: 1, sec: 'hdr', grp: 'h3' }),
      F('POSTEDON', 'تاريخ الترحيل', 'Posted on', 'text', { sys: 1, sec: 'hdr', grp: 'h3' })
    ]
  },

  gllines: {
    sheet: 'GLJournalLines', perm: 'gl', key: ['LINEID'], autoId: 'LINEID', title: { ar: 'سطور القيود', en: 'Journal entry lines' },
    sections: [{ id: 'l', ar: 'السطر', en: 'Line', groups: [{ id: 'l1', ar: 'الحساب', en: 'Account' }, { id: 'l2', ar: 'المبلغ', en: 'Amount' }] }],
    fields: [
      F('LINEID', 'معرّف السطر', 'Line ID', 'text', { key: 1, sys: 1, sec: 'l', grp: 'l1' }),
      F('VOUCHER', 'رقم القيد', 'Voucher', 'text', { req: 1, sec: 'l', grp: 'l1', list: 1 }),
      F('LINENUM', 'سطر القيد', 'Line number', 'number', { dec: 0, sec: 'l', grp: 'l1', list: 1 }),
      F('ACCOUNTTYPE', 'نوع الحساب', 'Account type', 'select', { opts: ACCOUNT_KINDS, sec: 'l', grp: 'l1', list: 1 }),
      F('PARTYID', 'رقم العميل / المورد / البنك', 'Party account', 'text', { sec: 'l', grp: 'l1', list: 1 }),
      F('MAINACCOUNTID', 'رقم الحساب', 'Main account', 'lookup', { lk: '@accounts', req: 1, sec: 'l', grp: 'l1', list: 1, aliases: ['Account', 'MAINACCOUNT'] }),
      F('ACCOUNTNAME', 'اسم الحساب', 'Account name', 'text', { calc: 1, sec: 'l', grp: 'l1', list: 1 })
    ].concat(unitFields_('', 'l', 'l1')).concat([
      F('CURRENCYCODE', 'العملة', 'Currency', 'lookup', { lk: '@currencies', sec: 'l', grp: 'l2', list: 1, aliases: ['Currency'] }),
      F('EXCHANGERATE', 'سعر الصرف', 'Exchange rate', 'number', { dec: 6, sec: 'l', grp: 'l2', list: 1 }),
      F('AMOUNTCURDEBIT', 'مدين (بالعملة)', 'Debit (currency)', 'number', { dec: 3, sec: 'l', grp: 'l2', list: 1 }),
      F('AMOUNTCURCREDIT', 'دائن (بالعملة)', 'Credit (currency)', 'number', { dec: 3, sec: 'l', grp: 'l2', list: 1 }),
      F('DEBITAMOUNT', 'مدين (محلي)', 'Debit (accounting)', 'number', { dec: 3, calc: 1, sec: 'l', grp: 'l2', list: 1 }),
      F('CREDITAMOUNT', 'دائن (محلي)', 'Credit (accounting)', 'number', { dec: 3, calc: 1, sec: 'l', grp: 'l2', list: 1 }),
      F('NOTE', 'ملاحظة', 'Note', 'text', { sec: 'l', grp: 'l2', list: 1, aliases: ['Text', 'TEXT'] }),
      F('REFERENCE', 'المرجع', 'Reference', 'text', { sec: 'l', grp: 'l2', list: 1 }),
      F('LEDGERDIMENSIONDISPLAYVALUE', 'الحساب والأبعاد', 'Ledger account display value', 'text', { calc: 1, sec: 'l', grp: 'l1', list: 1,
        aliases: ['ACCOUNTDISPLAYVALUE'] })
    ])
  },

  currencies: {
    sheet: 'Currencies', perm: 'masterdata', key: ['CURRENCYCODE'], title: { ar: 'العملات', en: 'Currencies' },
    sections: [{ id: 'g', ar: 'العملة', en: 'Currency', groups: [{ id: 'g', ar: 'العملة', en: 'Currency' }] }],
    fields: [
      F('CURRENCYCODE', 'رمز العملة', 'Currency code', 'text', { req: 1, key: 1, sec: 'g', grp: 'g', list: 1, aliases: ['Currency', 'CODE'] }),
      F('NAME', 'اسم العملة', 'Name', 'text', { req: 1, sec: 'g', grp: 'g', list: 1, aliases: ['Description', 'TXT'] }),
      F('SYMBOL', 'الرمز', 'Symbol', 'text', { sec: 'g', grp: 'g', list: 1 })
    ]
  },

  exrates: {
    sheet: 'ExchangeRates', perm: 'masterdata', key: ['CURRENCYCODE', 'FROMDATE'], title: { ar: 'أسعار الصرف', en: 'Exchange rates' },
    sections: [{ id: 'g', ar: 'سعر الصرف', en: 'Exchange rate', groups: [{ id: 'g', ar: 'السعر مقابل العملة المحلية', en: 'Rate against local currency' }] }],
    fields: [
      F('CURRENCYCODE', 'العملة', 'Currency', 'lookup', { lk: '@currencies', req: 1, key: 1, sec: 'g', grp: 'g', list: 1, aliases: ['From currency', 'FROMCURRENCY'] }),
      F('FROMDATE', 'ساري من تاريخ', 'Start date', 'date', { req: 1, key: 1, dflt: '@today', sec: 'g', grp: 'g', list: 1, aliases: ['Start date', 'STARTDATE'] }),
      F('RATE', 'السعر (عدد وحدات العملة المحلية لكل وحدة)', 'Rate (local units per 1)', 'number', { dec: 6, req: 1, sec: 'g', grp: 'g', list: 1,
        aliases: ['Exchange rate', 'EXCHANGERATE'] })
    ]
  },

  fiscal: {
    sheet: 'FiscalYears', perm: 'fiscal', key: ['YEAR'], title: { ar: 'السنوات المالية', en: 'Fiscal years' },
    sections: [{ id: 'g', ar: 'السنة المالية', en: 'Fiscal year', groups: [
      { id: 'y', ar: 'الفترة', en: 'Period' }, { id: 'c', ar: 'الإقفال', en: 'Closing' }] }],
    fields: [
      F('YEAR', 'السنة المالية', 'Fiscal year', 'text', { req: 1, key: 1, sec: 'g', grp: 'y', list: 1 }),
      F('STARTDATE', 'تاريخ البداية', 'Start date', 'date', { sec: 'g', grp: 'y', list: 1 }),
      F('ENDDATE', 'تاريخ النهاية', 'End date', 'date', { sec: 'g', grp: 'y', list: 1 }),
      F('RETAINEDEARNINGSACCOUNT', 'حساب الأرباح المحتجزة', 'Retained earnings account', 'lookup', { lk: '@accounts', sec: 'g', grp: 'c', list: 1 }),
      F('STATUS', 'الحالة', 'Status', 'select', { opts: FISCAL_STATUS, sys: 1, sec: 'g', grp: 'c', list: 1 }),
      F('CLOSINGVOUCHER', 'قيد الإقفال', 'Closing voucher', 'text', { sys: 1, sec: 'g', grp: 'c', list: 1 }),
      F('CLOSEDBY', 'أقفلها', 'Closed by', 'text', { sys: 1, sec: 'g', grp: 'c', list: 1 }),
      F('CLOSEDON', 'تاريخ الإقفال', 'Closed on', 'text', { sys: 1, sec: 'g', grp: 'c' })
    ]
  },

  customers: {
    sheet: 'Customers', perm: 'customers', key: ['CUSTOMERACCOUNT'], title: { ar: 'العملاء', en: 'Customers' },
    sections: [{ id: 'general', ar: 'عام', en: 'General', groups: [
      { id: 'id', ar: 'التعريف', en: 'Identification' }, { id: 'contact', ar: 'الاتصال', en: 'Contact' }, { id: 'acc', ar: 'المحاسبة', en: 'Accounting' }] }],
    fields: [
      F('CUSTOMERACCOUNT', 'رقم العميل', 'Customer account', 'text', { req: 1, key: 1, sec: 'general', grp: 'id', list: 1, aliases: ['Customer', 'ACCOUNTNUM'] }),
      F('NAME', 'اسم العميل', 'Name', 'text', { req: 1, sec: 'general', grp: 'id', list: 1, aliases: ['Customer name', 'ORGANIZATIONNAME'] }),
      F('ADDRESS', 'العنوان', 'Address', 'textarea', { sec: 'general', grp: 'contact', list: 1, aliases: ['FULLPRIMARYADDRESS', 'ADDRESSDESCRIPTION'] }),
      F('EMAIL', 'البريد الإلكتروني', 'Email', 'email', { sec: 'general', grp: 'contact', list: 1, aliases: ['PRIMARYCONTACTEMAIL', 'Primary email'] }),
      F('PHONE', 'رقم التليفون', 'Phone', 'text', { sec: 'general', grp: 'contact', list: 1, aliases: ['PRIMARYCONTACTPHONE', 'Telephone'] }),
      F('CURRENCYCODE', 'العملة', 'Currency', 'lookup', { lk: '@currencies', sec: 'general', grp: 'acc', list: 1, aliases: ['SALESCURRENCYCODE'] }),
      F('LEDGERACCOUNT', 'حساب العميل في الدليل', 'Customer ledger account', 'lookup', { lk: '@accounts', req: 1, sec: 'general', grp: 'acc', list: 1,
        aliases: ['Summary account', 'Receivable account'] }),
      F('OPENINGBALANCE', 'الرصيد الافتتاحي (مدين +)', 'Opening balance (debit +)', 'number', { dec: 3, sec: 'general', grp: 'acc', list: 1 }),
      F('OPENINGVOUCHER', 'قيد الرصيد الافتتاحي', 'Opening voucher', 'text', { sys: 1, sec: 'general', grp: 'acc' })
    ]
  },

  banks: {
    sheet: 'BankAccounts', perm: 'banks', key: ['BANKACCOUNTID'], title: { ar: 'الصناديق والبنوك', en: 'Cash & bank accounts' },
    sections: [{ id: 'general', ar: 'عام', en: 'General', groups: [
      { id: 'id', ar: 'التعريف', en: 'Identification' }, { id: 'acc', ar: 'المحاسبة', en: 'Accounting' }] }],
    fields: [
      F('BANKACCOUNTID', 'رقم الصندوق / البنك', 'Cash/bank account', 'text', { req: 1, key: 1, sec: 'general', grp: 'id', list: 1, aliases: ['Bank account', 'ACCOUNTID'] }),
      F('NAME', 'اسم الصندوق / البنك', 'Name', 'text', { req: 1, sec: 'general', grp: 'id', list: 1, aliases: ['Bank name'] }),
      F('BANKGROUP', 'المجموعة', 'Group', 'lookup', { lk: 'BankGroup', sec: 'general', grp: 'id', list: 1, aliases: ['BANKGROUPID', 'Group'] }),
      F('CURRENCYCODE', 'العملة', 'Currency', 'lookup', { lk: '@currencies', sec: 'general', grp: 'acc', list: 1 }),
      F('LEDGERACCOUNT', 'الحساب في الدليل', 'Ledger account', 'lookup', { lk: '@accounts', req: 1, sec: 'general', grp: 'acc', list: 1,
        aliases: ['Main account', 'MAINACCOUNT'] }),
      F('OPENINGBALANCE', 'الرصيد الافتتاحي (مدين +)', 'Opening balance (debit +)', 'number', { dec: 3, sec: 'general', grp: 'acc', list: 1 }),
      F('OPENINGVOUCHER', 'قيد الرصيد الافتتاحي', 'Opening voucher', 'text', { sys: 1, sec: 'general', grp: 'acc' })
    ]
  },

  itemgroups: {
    sheet: 'ItemGroups', perm: 'items', key: ['ITEMGROUPID'], title: { ar: 'مجموعات الأصناف', en: 'Item groups' },
    sections: [{ id: 'g', ar: 'المجموعة', en: 'Group', groups: [{ id: 'g1', ar: 'التعريف', en: 'Identification' }, { id: 'g2', ar: 'الربط المحاسبي', en: 'Posting' }] }],
    fields: [
      F('ITEMGROUPID', 'رقم المجموعة', 'Item group', 'text', { req: 1, key: 1, sec: 'g', grp: 'g1', list: 1, aliases: ['ITEMGROUP', 'Item group'] }),
      F('NAME', 'اسم المجموعة', 'Name', 'text', { req: 1, sec: 'g', grp: 'g1', list: 1 }),
      F('SALESACCOUNT', 'حساب المبيعات (الإيراد)', 'Sales revenue account', 'lookup', { lk: '@accounts', sec: 'g', grp: 'g2', list: 1 }),
      F('PURCHASEACCOUNT', 'حساب المشتريات (المصروف)', 'Purchase / expense account', 'lookup', { lk: '@accounts', sec: 'g', grp: 'g2', list: 1 }),
      F('SALESRETURNACCOUNT', 'حساب مردودات المبيعات', 'Sales returns account', 'lookup', { lk: '@accounts', sec: 'g', grp: 'g2', list: 1 }),
      F('PURCHRETURNACCOUNT', 'حساب مردودات المشتريات', 'Purchase returns account', 'lookup', { lk: '@accounts', sec: 'g', grp: 'g2', list: 1 })
    ]
  },

  salesitems: {
    sheet: 'SalesItems', perm: 'items', key: ['ITEMID'], title: { ar: 'أصناف البيع', en: 'Sales items' },
    sections: [{ id: 'i', ar: 'الصنف', en: 'Item', groups: [{ id: 'i1', ar: 'التعريف', en: 'Identification' }, { id: 'i2', ar: 'السعر', en: 'Price' }] }],
    fields: [
      F('ITEMID', 'رقم الصنف', 'Item number', 'text', { req: 1, key: 1, sec: 'i', grp: 'i1', list: 1, aliases: ['ITEMNUMBER', 'Item number'] }),
      F('NAME', 'اسم الصنف', 'Item name', 'text', { req: 1, sec: 'i', grp: 'i1', list: 1, aliases: ['PRODUCTNAME', 'Product name'] }),
      F('ITEMGROUP', 'مجموعة الصنف', 'Item group', 'lookup', { lk: '@itemgroups', req: 1, sec: 'i', grp: 'i1', list: 1, aliases: ['ITEMGROUPID'] }),
      F('UNIT', 'الوحدة', 'Unit', 'lookup', { lk: 'ItemUnit', sec: 'i', grp: 'i2', list: 1, aliases: ['SALESUNITSYMBOL', 'Unit'] }),
      F('PRICE', 'سعر البيع', 'Sales price', 'number', { dec: 3, sec: 'i', grp: 'i2', list: 1, aliases: ['SALESPRICE', 'Price'] })
    ]
  },

  purchitems: {
    sheet: 'PurchaseItems', perm: 'items', key: ['ITEMID'], title: { ar: 'أصناف الشراء', en: 'Purchase items' },
    sections: [{ id: 'i', ar: 'الصنف', en: 'Item', groups: [{ id: 'i1', ar: 'التعريف', en: 'Identification' }, { id: 'i2', ar: 'السعر', en: 'Price' }] }],
    fields: [
      F('ITEMID', 'رقم الصنف', 'Item number', 'text', { req: 1, key: 1, sec: 'i', grp: 'i1', list: 1, aliases: ['ITEMNUMBER', 'Item number'] }),
      F('NAME', 'اسم الصنف', 'Item name', 'text', { req: 1, sec: 'i', grp: 'i1', list: 1, aliases: ['PRODUCTNAME', 'Product name'] }),
      F('ITEMGROUP', 'مجموعة الصنف', 'Item group', 'lookup', { lk: '@itemgroups', req: 1, sec: 'i', grp: 'i1', list: 1, aliases: ['ITEMGROUPID'] }),
      F('UNIT', 'الوحدة', 'Unit', 'lookup', { lk: 'ItemUnit', sec: 'i', grp: 'i2', list: 1, aliases: ['PURCHASEUNITSYMBOL', 'Unit'] }),
      F('PRICE', 'سعر الشراء', 'Purchase price', 'number', { dec: 3, sec: 'i', grp: 'i2', list: 1, aliases: ['PURCHASEPRICE', 'Price'] })
    ]
  },

  invoices: {
    sheet: 'Invoices', perm: 'invoices', key: ['INVOICEID'], title: { ar: 'فواتير المبيعات والمشتريات', en: 'Sales & purchase invoices' },
    sections: [{ id: 'h', ar: 'رأس الفاتورة', en: 'Invoice header', groups: [
      { id: 'h1', ar: 'الفاتورة', en: 'Invoice' }, { id: 'h2', ar: 'الطرف', en: 'Party' }, { id: 'h3', ar: 'المبالغ', en: 'Amounts' }, { id: 'h4', ar: 'الترحيل', en: 'Posting' }] }],
    fields: [
      F('INVOICEID', 'رقم الفاتورة', 'Invoice', 'text', { key: 1, sys: 1, sec: 'h', grp: 'h1', list: 1 }),
      F('INVOICETYPE', 'نوع الفاتورة', 'Invoice type', 'select', { opts: INVOICE_TYPES, req: 1, sec: 'h', grp: 'h1', list: 1 }),
      F('INVOICEDATE', 'تاريخ الفاتورة', 'Invoice date', 'date', { req: 1, dflt: '@today', sec: 'h', grp: 'h1', list: 1 }),
      F('DUEDATE', 'تاريخ الاستحقاق', 'Due date', 'date', { sec: 'h', grp: 'h1', list: 1 }),
      F('PARTYID', 'العميل / المورد', 'Customer / vendor', 'lookup', { lkBy: 'INVOICETYPE', lkMap: { Sales: '@customers', Purchase: '@vendors', SalesCredit: '@customers', SalesDebit: '@customers', PurchaseDebit: '@vendors', PurchaseCredit: '@vendors', Ledger: '@customers' }, req: 1, sec: 'h', grp: 'h2', list: 1 }),
      F('PARTYNAME', 'اسم العميل / المورد', 'Party name', 'text', { calc: 1, sec: 'h', grp: 'h2', list: 1 }),
      F('ORIGINVOICE', 'الفاتورة الأصلية', 'Original invoice', 'lookup', { lk: '@noteinvoices', sec: 'h', grp: 'h2', list: 1 }),
      F('PARTYREF', 'مرجع الطرف (رقم فاتورة المورد / أمر الشراء)', 'Party reference', 'text', { sec: 'h', grp: 'h2', list: 1 }),
      F('SUB2_OPERATINGUNIT', 'الوحدة الفرعية 2', 'A_Sub2_OperatingUnit', 'lookup', { lk: 'Sub2_OperatingUnit', up: 'SUB1_OPERATINGUNIT', req: 1, sec: 'h', grp: 'h1', list: 1 }),
      F('SUB1_OPERATINGUNIT', 'الوحدة الفرعية 1', 'A_Sub1_OperatingUnit', 'lookup', { lk: 'Sub1_OperatingUnit', up: 'OPERATINGUNIT', calc: 1, sec: 'h', grp: 'h1' }),
      F('OPERATINGUNIT', 'الوحدة الرئيسية', 'A_A_OperatingUnit', 'lookup', { lk: 'OperatingUnit', calc: 1, sec: 'h', grp: 'h1' }),
      F('DESCRIPTION', 'البيان', 'Description', 'text', { sec: 'h', grp: 'h2', list: 1 }),
      F('CURRENCYCODE', 'العملة', 'Currency', 'text', { calc: 1, sec: 'h', grp: 'h3', list: 1 }),
      F('EXCHANGERATE', 'سعر الصرف', 'Exchange rate', 'number', { dec: 6, sec: 'h', grp: 'h3' }),
      F('SUBTOTAL', 'الإجمالي قبل الضريبة', 'Subtotal', 'number', { dec: 3, sys: 1, sec: 'h', grp: 'h3', list: 1 }),
      F('TAXAMOUNT', 'الضريبة', 'Tax', 'number', { dec: 3, sys: 1, sec: 'h', grp: 'h3', list: 1 }),
      F('TOTALAMOUNT', 'إجمالي الفاتورة', 'Invoice total', 'number', { dec: 3, sys: 1, sec: 'h', grp: 'h3', list: 1 }),
      F('PAIDAMOUNT', 'المسدد', 'Paid', 'number', { dec: 3, sys: 1, sec: 'h', grp: 'h3', list: 1 }),
      F('CREDITAMOUNT', 'المخصوم بإشعارات', 'Credited', 'number', { dec: 3, sys: 1, sec: 'h', grp: 'h3', list: 1 }),
      F('STATUS', 'الحالة', 'Status', 'select', { opts: CASH_STATUS, sys: 1, sec: 'h', grp: 'h4', list: 1 }),
      F('GLVOUCHER', 'رقم القيد', 'Journal voucher', 'text', { sys: 1, sec: 'h', grp: 'h4', list: 1 }),
      F('REVERSALVOUCHER', 'قيد الإلغاء', 'Reversal voucher', 'text', { sys: 1, sec: 'h', grp: 'h4' }),
      F('POSTEDBY', 'رحّل بواسطة', 'Posted by', 'text', { sys: 1, sec: 'h', grp: 'h4' }),
      F('POSTEDON', 'تاريخ الترحيل', 'Posted on', 'text', { sys: 1, sec: 'h', grp: 'h4' })
    ]
  },

  invlines: {
    sheet: 'InvoiceLines', perm: 'invoices', key: ['LINEID'], autoId: 'LINEID', title: { ar: 'سطور الفواتير', en: 'Invoice lines' },
    sections: [{ id: 'l', ar: 'السطر', en: 'Line', groups: [{ id: 'l1', ar: 'الصنف', en: 'Item' }, { id: 'l2', ar: 'المبلغ', en: 'Amount' }] }],
    fields: [
      F('LINEID', 'معرّف السطر', 'Line ID', 'text', { key: 1, sys: 1, sec: 'l', grp: 'l1' }),
      F('INVOICEID', 'رقم الفاتورة', 'Invoice', 'text', { req: 1, sec: 'l', grp: 'l1', list: 1 }),
      F('LINENUM', 'رقم السطر', 'Line number', 'number', { dec: 0, sec: 'l', grp: 'l1', list: 1 }),
      F('ITEMID', 'رقم الصنف', 'Item number', 'text', { sec: 'l', grp: 'l1', list: 1 }),
      F('ITEMDESCRIPTION', 'البيان / الصنف', 'Item / description', 'text', { req: 1, sec: 'l', grp: 'l1', list: 1 }),
      F('UNIT', 'الوحدة', 'Unit', 'text', { sec: 'l', grp: 'l1', list: 1 }),
      F('MAINACCOUNTID', 'رقم الحساب', 'Main account', 'lookup', { lk: '@accounts', req: 1, sec: 'l', grp: 'l1', list: 1 }),
      F('ACCOUNTNAME', 'اسم الحساب', 'Account name', 'text', { calc: 1, sec: 'l', grp: 'l1', list: 1 })
    ].concat(unitFields_('', 'l', 'l1')).concat([
      F('QTY', 'الكمية', 'Quantity', 'number', { dec: 3, req: 1, sec: 'l', grp: 'l2', list: 1 }),
      F('UNITPRICE', 'سعر الوحدة', 'Unit price', 'number', { dec: 3, req: 1, sec: 'l', grp: 'l2', list: 1 }),
      F('LINEAMOUNT', 'الإجمالي', 'Line amount', 'number', { dec: 3, calc: 1, sec: 'l', grp: 'l2', list: 1 }),
      F('TAXPERCENT', 'نسبة الضريبة %', 'Tax %', 'number', { dec: 2, sec: 'l', grp: 'l2', list: 1 }),
      F('TAXAMOUNT', 'الضريبة', 'Tax', 'number', { dec: 3, calc: 1, sec: 'l', grp: 'l2', list: 1 })
    ])
  },

  budgetheads: {
    sheet: 'BudgetHeaders', perm: 'budget', key: ['BUDGETID'], title: { ar: 'الموازنات التقديرية', en: 'Budgets' },
    sections: [{ id: 'h', ar: 'رأس الموازنة', en: 'Budget header', groups: [
      { id: 'h1', ar: 'الموازنة', en: 'Budget' }, { id: 'h2', ar: 'البيان', en: 'Description' }, { id: 'h3', ar: 'الاعتماد', en: 'Approval' }] }],
    fields: [
      F('BUDGETID', 'رقم الموازنة', 'Budget ID', 'text', { key: 1, sys: 1, sec: 'h', grp: 'h1', list: 1 }),
      F('BUDGETYEAR', 'سنة الموازنة', 'Budget year', 'text', { req: 1, sec: 'h', grp: 'h1', list: 1, aliases: ['Year'] }),
      F('BUDGETTYPE', 'نوع الموازنة', 'Budget type', 'select', { opts: BUDGET_TYPES, req: 1, dflt: 'Expense', sec: 'h', grp: 'h1', list: 1 }),
      F('ENTRYDATE', 'تاريخ الإدخال', 'Entry date', 'date', { req: 1, dflt: '@today', sec: 'h', grp: 'h1', list: 1 }),
      F('PERIODTYPE', 'سنوي / شهري', 'Annual / monthly', 'select', { opts: BUDGET_PERIODS, req: 1, dflt: 'Annual', sec: 'h', grp: 'h1', list: 1 }),
      F('DESCRIPTION', 'البيان', 'Description', 'text', { sec: 'h', grp: 'h2', list: 1 }),
      F('TOTALAMOUNT', 'إجمالي المبلغ', 'Total amount', 'number', { dec: 3, sys: 1, sec: 'h', grp: 'h3', list: 1 }),
      F('STATUS', 'الحالة', 'Status', 'select', { opts: BUDGET_STATUS, sys: 1, sec: 'h', grp: 'h3', list: 1 }),
      F('APPROVEDBY', 'اعتمد بواسطة', 'Approved by', 'text', { sys: 1, sec: 'h', grp: 'h3', list: 1 }),
      F('APPROVEDON', 'تاريخ الاعتماد', 'Approved on', 'text', { sys: 1, sec: 'h', grp: 'h3' }),
      F('POSTEDBY', 'رحّل بواسطة', 'Posted by', 'text', { sys: 1, sec: 'h', grp: 'h3', list: 1 }),
      F('POSTEDON', 'تاريخ الترحيل', 'Posted on', 'text', { sys: 1, sec: 'h', grp: 'h3' })
    ]
  },

  cashvouchers: {
    sheet: 'CashVouchers', perm: 'cashvouchers', key: ['VOUCHERNO'], title: { ar: 'سندات القبض والصرف', en: 'Receipts & payments' },
    sections: [{ id: 'v', ar: 'السند', en: 'Voucher', groups: [
      { id: 'v1', ar: 'السند', en: 'Voucher' }, { id: 'v2', ar: 'الطرف', en: 'Party' }, { id: 'v3', ar: 'المبلغ', en: 'Amount' }, { id: 'v4', ar: 'الترحيل', en: 'Posting' }] }],
    fields: [
      F('VOUCHERNO', 'رقم السند', 'Voucher number', 'text', { key: 1, sys: 1, sec: 'v', grp: 'v1', list: 1 }),
      F('VOUCHERTYPE', 'نوع السند', 'Voucher type', 'select', { opts: CASH_TYPES, req: 1, sec: 'v', grp: 'v1', list: 1 }),
      F('TRANSDATE', 'التاريخ', 'Date', 'date', { req: 1, dflt: '@today', sec: 'v', grp: 'v1', list: 1 }),
      F('BANKACCOUNTID', 'الصندوق / البنك', 'Cash/bank account', 'lookup', { lk: '@banks', req: 1, sec: 'v', grp: 'v1', list: 1 }),
      F('PARTYTYPE', 'نوع الطرف', 'Party type', 'select', { opts: CASH_PARTY, req: 1, sec: 'v', grp: 'v2', list: 1 }),
      F('PARTYID', 'العميل / المورد / الحساب', 'Party', 'lookup', { lkBy: 'PARTYTYPE', lkMap: { Customer: '@customers', Vendor: '@vendors', Ledger: '@accounts' }, req: 1, sec: 'v', grp: 'v2', list: 1 }),
      F('PARTYNAME', 'اسم الطرف', 'Party name', 'text', { calc: 1, sec: 'v', grp: 'v2', list: 1 }),
      F('SUB2_OPERATINGUNIT', 'الوحدة الفرعية 2', 'A_Sub2_OperatingUnit', 'lookup', { lk: 'Sub2_OperatingUnit', sec: 'v', grp: 'v2' }),
      F('INVOICEID', 'سداد الفاتورة', 'Settles invoice', 'lookup', { lk: '@openinvoices', sec: 'v', grp: 'v2', list: 1 }),
      F('AMOUNT', 'المبلغ', 'Amount', 'number', { dec: 3, req: 1, sec: 'v', grp: 'v3', list: 1 }),
      F('CURRENCYCODE', 'العملة', 'Currency', 'text', { calc: 1, sec: 'v', grp: 'v3', list: 1 }),
      F('EXCHANGERATE', 'سعر الصرف', 'Exchange rate', 'number', { dec: 6, sec: 'v', grp: 'v3' }),
      F('PAYMENTMETHOD', 'طريقة الدفع', 'Payment method', 'select', { opts: PAY_METHODS, dflt: 'Cash', sec: 'v', grp: 'v3', list: 1 }),
      F('CHEQUENUMBER', 'رقم الشيك / الحوالة', 'Cheque / transfer no.', 'text', { sec: 'v', grp: 'v3', list: 1 }),
      F('DESCRIPTION', 'البيان', 'Description', 'text', { sec: 'v', grp: 'v3', list: 1 }),
      F('STATUS', 'الحالة', 'Status', 'select', { opts: CASH_STATUS, sys: 1, sec: 'v', grp: 'v4', list: 1 }),
      F('GLVOUCHER', 'رقم القيد', 'Journal voucher', 'text', { sys: 1, sec: 'v', grp: 'v4', list: 1 }),
      F('REVERSALVOUCHER', 'قيد الإلغاء', 'Reversal voucher', 'text', { sys: 1, sec: 'v', grp: 'v4' }),
      F('POSTEDBY', 'رحّل بواسطة', 'Posted by', 'text', { sys: 1, sec: 'v', grp: 'v4' }),
      F('POSTEDON', 'تاريخ الترحيل', 'Posted on', 'text', { sys: 1, sec: 'v', grp: 'v4' })
    ]
  },

  budgets: {
    sheet: 'Budgets', perm: 'budget', key: ['LINEID'], autoId: 'LINEID', title: { ar: 'سطور الموازنة التقديرية', en: 'Budget lines' },
    sections: [{ id: 'b', ar: 'سطر الموازنة', en: 'Budget line', groups: [
      { id: 'a', ar: 'الحساب', en: 'Account' }, { id: 'u', ar: 'الوحدة (سجّل الفرعية 2)', en: 'Unit (enter sub-unit 2)' }, { id: 'm', ar: 'المبلغ', en: 'Amount' }] }],
    fields: [
      F('LINEID', 'معرّف السطر', 'Line ID', 'text', { key: 1, sys: 1, sec: 'b', grp: 'm' }),
      F('BUDGETID', 'رقم الموازنة', 'Budget ID', 'text', { req: 1, sec: 'b', grp: 'a', list: 1 }),
      F('LINENUM', 'رقم السطر', 'Line number', 'number', { dec: 0, sec: 'b', grp: 'a', list: 1 }),
      F('BUDGETYEAR', 'سنة الموازنة', 'Budget year', 'text', { calc: 1, sec: 'b', grp: 'a', list: 1, aliases: ['Year', 'YEAR'] }),
      F('BUDGETTYPE', 'نوع الموازنة', 'Budget type', 'select', { opts: BUDGET_TYPES, calc: 1, sec: 'b', grp: 'a', list: 1 }),
      F('MAINACCOUNTID', 'رقم الحساب', 'Main account', 'lookup', { lk: '@accounts', req: 1, sec: 'b', grp: 'a', list: 1, aliases: ['Account', 'MAINACCOUNT'] }),
      F('ACCOUNTNAME', 'اسم الحساب', 'Account name', 'text', { calc: 1, sec: 'b', grp: 'a', list: 1 })
    ].concat(unitFields_('', 'b', 'u', { ou: { list: 1 } })).concat([
      F('AMOUNT', 'المبلغ التقديري', 'Budget amount', 'number', { dec: 3, req: 1, sec: 'b', grp: 'm', list: 1, aliases: ['Amount', 'Budget'] }),
      F('PERIODTYPE', 'سنوي / شهري', 'Annual / monthly', 'select', { opts: BUDGET_PERIODS, calc: 1, sec: 'b', grp: 'm', list: 1, aliases: ['Period'] }),
      F('NOTE', 'ملاحظة', 'Note', 'text', { sec: 'b', grp: 'm', list: 1 })
    ])
  },

  vendors: {
    sheet: 'Vendors', perm: 'vendors', key: ['VENDORACCOUNTNUMBER'], title: { ar: 'الموردون', en: 'Vendors' },
    sections: [{ id: 'general', ar: 'عام', en: 'General', groups: [
      { id: 'id', ar: 'التعريف', en: 'Identification' }, { id: 'contact', ar: 'الاتصال', en: 'Contact' }, { id: 'acc', ar: 'المحاسبة', en: 'Accounting' }] }],
    fields: [
      F('VENDORACCOUNTNUMBER', 'رقم المورد', 'Vendor account', 'text', { req: 1, key: 1, sec: 'general', grp: 'id', list: 1,
        aliases: ['Vendor account number', 'Vendor', 'ACCOUNTNUM'] }),
      F('VENDORORGANIZATIONNAME', 'اسم المورد', 'Name', 'text', { req: 1, sec: 'general', grp: 'id', list: 1,
        aliases: ['Vendor name', 'NAME', 'ORGANIZATIONNAME'] }),
      F('ADDRESS', 'العنوان', 'Address', 'textarea', { sec: 'general', grp: 'contact', list: 1,
        aliases: ['FULLPRIMARYADDRESS', 'Primary address', 'ADDRESSDESCRIPTION', 'ADDRESSSTREET'] }),
      F('PRIMARYEMAILADDRESS', 'البريد الإلكتروني', 'Email', 'email', { sec: 'general', grp: 'contact', list: 1,
        aliases: ['Primary email', 'EMAIL', 'Email address'] }),
      F('VENDORLEDGERACCOUNT', 'حساب المورد في الدليل', 'Vendor ledger account', 'lookup', { lk: '@accounts', sec: 'general', grp: 'acc', list: 1,
        aliases: ['Summary account', 'Ledger account', 'Payable account'] }),
      F('CURRENCYCODE', 'العملة', 'Currency', 'lookup', { lk: '@currencies', sec: 'general', grp: 'acc', list: 1 }),
      F('OPENINGBALANCE', 'الرصيد الافتتاحي (دائن +)', 'Opening balance (credit +)', 'number', { dec: 3, sec: 'general', grp: 'acc', list: 1 }),
      F('OPENINGVOUCHER', 'قيد الرصيد الافتتاحي', 'Opening voucher', 'text', { sys: 1, sec: 'general', grp: 'acc' })
    ]
  },

  groups: {
    sheet: 'AssetGroups', perm: 'masterdata', key: ['FIXEDASSETGROUPID'], title: { ar: 'مجموعات الأصول', en: 'Fixed asset groups' },
    sections: [
      { id: 'general', ar: 'عام', en: 'General', groups: [{ id: 'id', ar: 'المجموعة', en: 'Group' }] },
      { id: 'acc', ar: 'حسابات الترحيل (من دليل الحسابات)', en: 'Posting accounts (chart of accounts)', groups: [
        { id: 'asset', ar: 'حسابات الأصل', en: 'Asset accounts' }, { id: 'pl', ar: 'حسابات المصروف والاستبعاد', en: 'Expense & disposal accounts' }] }
    ],
    fields: [
      F('FIXEDASSETGROUPID', 'كود المجموعة', 'Fixed asset group', 'text', { req: 1, key: 1, sec: 'general', grp: 'id', list: 1,
        aliases: ['Group', 'Fixed asset group ID', 'GROUPID', 'CODE'] }),
      F('NAME', 'اسم المجموعة', 'Name', 'text', { req: 1, sec: 'general', grp: 'id', list: 1, aliases: ['Description'] }),
      F('ASSETACCOUNT', 'حساب الأصل (التكلفة / الاقتناء)', 'Fixed asset account (acquisition)', 'lookup', { lk: '@accounts', sec: 'acc', grp: 'asset', list: 1 }),
      F('ACCUMDEPACCOUNT', 'حساب مجمع الإهلاك', 'Accumulated depreciation account', 'lookup', { lk: '@accounts', sec: 'acc', grp: 'asset', list: 1 }),
      F('DEPEXPENSEACCOUNT', 'حساب مصروف الإهلاك', 'Depreciation expense account', 'lookup', { lk: '@accounts', sec: 'acc', grp: 'pl', list: 1 }),
      F('DISPOSALACCOUNT', 'حساب أرباح / خسائر الاستبعاد', 'Disposal gain/loss account', 'lookup', { lk: '@accounts', sec: 'acc', grp: 'pl', list: 1 })
    ]
  },

  accounts: {
    sheet: 'MainAccounts', perm: 'masterdata', key: ['MAINACCOUNTID'], title: { ar: 'دليل الحسابات', en: 'Chart of accounts' },
    sections: [{ id: 'general', ar: 'عام', en: 'General', groups: [
      { id: 'id', ar: 'الحساب', en: 'Account' }, { id: 'tree', ar: 'الشجرة', en: 'Hierarchy' }, { id: 'cls', ar: 'التصنيف', en: 'Classification' }] }],
    fields: [
      F('MAINACCOUNTID', 'رقم الحساب', 'Account number', 'text', { req: 1, key: 1, sec: 'general', grp: 'id', list: 1,
        aliases: ['Main account', 'Main account ID', 'MAINACCOUNT', 'Account', 'AccountNumber'] }),
      F('NAME', 'اسم الحساب', 'Account name', 'text', { req: 1, sec: 'general', grp: 'id', list: 1, aliases: ['Name', 'AccountName'] }),
      F('ACCOUNTLEVEL', 'مستوى الحساب', 'Account level', 'select', { opts: ACCOUNT_LEVELS, req: 1, dflt: 'Posting', sec: 'general', grp: 'tree', list: 1,
        aliases: ['Level'] }),
      F('PARENTACCOUNT', 'الحساب الأب (الرئيسي)', 'Parent (header) account', 'lookup', { lk: '@accountHeaders', sec: 'general', grp: 'tree', list: 1,
        aliases: ['Parent account', 'Parent'] }),
      F('MAINACCOUNTCATEGORY', 'مجموعة الحساب الرئيسي', 'Main account group', 'text', { sec: 'general', grp: 'cls', list: 1,
        aliases: ['Main account category', 'MainAccountGroup', 'Account category', 'MAINACCOUNTCATEGORYID'] }),
      F('CASHFLOWCATEGORY', 'تصنيف التدفق النقدي', 'Cash flow category', 'select', { opts: CASHFLOW_CATS, sec: 'general', grp: 'cls',
        aliases: ['Cash flow'] }),
      F('MAINACCOUNTTYPE', 'نوع الحساب', 'Account type', 'select', { opts: ACCOUNT_TYPES, req: 1, sec: 'general', grp: 'cls', list: 1,
        aliases: ['Main account type', 'Type', 'AccountType'] })
    ]
  }
};

// إنشاء مخطط لكل قائمة أساسية (شيت مستقل لكل نوع)
LK.forEach(function (t) {
  const fields = [
    F('CODE', 'الكود', 'ID', 'text', { req: 1, key: 1, sec: 'general', grp: 'g', list: 1,
      aliases: ['Value', 'ID', 'Code', t.d365 || t.v, 'DIMENSIONVALUE'] }),
    F('NAME', 'الاسم', 'Name', 'text', { req: 1, sec: 'general', grp: 'g', list: 1, aliases: ['Description'] })
  ];
  if (t.parent) {
    fields.push(F('PARENT', 'يتبع: ' + LKMAP[t.parent].ar, 'Parent: ' + LKMAP[t.parent].en, 'lookup', { lk: t.parent, req: t.req ? 1 : 0,
      sec: 'general', grp: 'g', list: 1, aliases: ['Parent', 'Parent code', LKMAP[t.parent].d365 || t.parent] }));
  }
  SCHEMAS[t.v] = {
    sheet: t.sheet, perm: 'masterdata', key: ['CODE'], title: { ar: t.ar, en: t.en }, lkType: t.v, parentType: t.parent || '',
    sections: [{ id: 'general', ar: 'عام', en: 'General', groups: [{ id: 'g', ar: 'القيمة', en: 'Value' }] }],
    fields: fields
  };
});

/* ============================== Web app ============================== */

function doGet() {
  ensureSetup_();
  const t = HtmlService.createTemplateFromFile('Index');
  return t.evaluate()
    .setTitle(getSettings_().ORGNAME || 'الإدارة المالية')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include_(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* ============================== Setup ============================== */

function getSs_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (ss) return ss;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(SS_PROP);
  if (id) return SpreadsheetApp.openById(id);
  ss = SpreadsheetApp.create('Fixed Assets DB');
  props.setProperty(SS_PROP, ss.getId());
  return ss;
}

function ensureSetup_() {
  if (PropertiesService.getScriptProperties().getProperty('FA_SETUP') === SETUP_VERSION) return;
  setup();
}

/** شغّل هذه الدالة مرة واحدة من محرر السكريبت لإنشاء الشيتات والمستخدم الافتراضي. */
function setup() {
  withLock_(function () {
    const ss = getSs_();
    Object.keys(SCHEMAS).forEach(function (e) {
      const sch = SCHEMAS[e];
      ensureSheet_(ss, sch.sheet, cols_(e), sch.fields.filter(function (f) { return f.type === 'number'; }).map(function (f) { return f.k; }));
    });
    ensureSheet_(ss, 'Users', USER_COLS, []);
    ensureSheet_(ss, 'Sessions', SESSION_COLS, []);
    ensureSheet_(ss, 'Settings', ['KEY', 'VALUE'], []);

    const st = readTable_('Settings');
    const have = {};
    st.rows.forEach(function (r) { have[r.KEY] = 1; });
    Object.keys(SETTING_DEFAULTS).forEach(function (k) {
      if (!have[k]) appendRow_(st, { KEY: k, VALUE: SETTING_DEFAULTS[k] });
    });

    const ut = readTable_('Users');
    if (!ut.rows.some(function (r) { return r.USERNAME; })) {
      const salt = Utilities.getUuid();
      appendRow_(ut, {
        USERNAME: 'admin', FULLNAME: 'مدير عام', EMAIL: '', PASSWORDHASH: hash_('Admin@123', salt), SALT: salt,
        ROLE: 'SuperAdmin', PERMISSIONS: '', STATUS: 'Active', MUSTCHANGE: 'TRUE', LASTLOGIN: '',
        CREATEDBY: 'setup', CREATEDON: stamp_(), MODIFIEDBY: '', MODIFIEDON: ''
      });
    }

    migrateV3_(ss);
    migrateV6_();
    migrateV7_();
    migrateV8_();
    migrateV9_();
    migrateV10_();
    migrateV11_();
    migrateV12_();
    migrateV13_();

    const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('ورقة1');
    if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
    PropertiesService.getScriptProperties().setProperty('FA_SETUP', SETUP_VERSION);
  });
  CacheService.getScriptCache().remove('SETTINGS');
}

/* ============================== Default data & migration ============================== */

// [رقم، اسم، المستوى، الأب، النوع، المجموعة]
const DEFAULT_ACCOUNTS = [
  ['1100', 'النقدية والبنوك', 'Header', '', 'Asset', 'CASH'], ['110100', 'البنك', 'Posting', '1100', 'Asset', 'CASH'],
  ['1200', 'الأصول الثابتة', 'Header', '', 'Asset', 'FA'],
  ['120100', 'المباني', 'Posting', '1200', 'Asset', 'FA'], ['120200', 'السيارات', 'Posting', '1200', 'Asset', 'FA'],
  ['120300', 'الأجهزة الطبية', 'Posting', '1200', 'Asset', 'FA'], ['120400', 'الأثاث', 'Posting', '1200', 'Asset', 'FA'],
  ['120500', 'أجهزة الحاسب', 'Posting', '1200', 'Asset', 'FA'],
  ['1290', 'مجمع إهلاك الأصول الثابتة', 'Header', '', 'Asset', 'FA'],
  ['129100', 'مجمع إهلاك المباني', 'Posting', '1290', 'Asset', 'FA'], ['129200', 'مجمع إهلاك السيارات', 'Posting', '1290', 'Asset', 'FA'],
  ['129300', 'مجمع إهلاك الأجهزة الطبية', 'Posting', '1290', 'Asset', 'FA'], ['129400', 'مجمع إهلاك الأثاث', 'Posting', '1290', 'Asset', 'FA'],
  ['129500', 'مجمع إهلاك أجهزة الحاسب', 'Posting', '1290', 'Asset', 'FA'],
  ['2100', 'الموردون', 'Header', '', 'Liability', 'AP'], ['210100', 'ذمم الموردين', 'Posting', '2100', 'Liability', 'AP'],
  ['3100', 'حقوق الملكية', 'Header', '', 'Equity', 'EQ'], ['310100', 'رأس المال', 'Posting', '3100', 'Equity', 'EQ'],
  ['4100', 'الإيرادات', 'Header', '', 'Revenue', 'REV'], ['410100', 'إيرادات الخدمات', 'Posting', '4100', 'Revenue', 'REV'],
  ['6100', 'مصروفات الإهلاك', 'Header', '', 'Expense', 'DEP'],
  ['610100', 'مصروف إهلاك المباني', 'Posting', '6100', 'Expense', 'DEP'], ['610200', 'مصروف إهلاك السيارات', 'Posting', '6100', 'Expense', 'DEP'],
  ['610300', 'مصروف إهلاك الأجهزة الطبية', 'Posting', '6100', 'Expense', 'DEP'], ['610400', 'مصروف إهلاك الأثاث', 'Posting', '6100', 'Expense', 'DEP'],
  ['610500', 'مصروف إهلاك أجهزة الحاسب', 'Posting', '6100', 'Expense', 'DEP'],
  ['7100', 'أرباح وخسائر الأصول', 'Header', '', 'ProfitAndLoss', 'GL'], ['710100', 'أرباح / خسائر استبعاد الأصول', 'Posting', '7100', 'ProfitAndLoss', 'GL']
];
// [كود، اسم، حساب الأصل، مجمع الإهلاك، مصروف الإهلاك، الاستبعاد]
const DEFAULT_GROUPS = [
  ['BLD', 'المباني', '120100', '129100', '610100', '710100'], ['VEH', 'السيارات', '120200', '129200', '610200', '710100'],
  ['MED', 'الأجهزة الطبية', '120300', '129300', '610300', '710100'], ['FUR', 'الأثاث', '120400', '129400', '610400', '710100'],
  ['IT', 'أجهزة الحاسب', '120500', '129500', '610500', '710100']
];
const DEFAULT_LK = {
  OperatingUnit: [['HE01', 'مستشفيات']], Sub1_OperatingUnit: [['DJI_HE0101', 'مستشفى جيبوتي', 'HE01']],
  Sub2_OperatingUnit: [['DJI_HE010001', 'عام المستشفى', 'DJI_HE0101']],
  Location: [['MAIN', 'المقر الرئيسي']], Building: [['B01', 'المبنى الإداري', 'MAIN']], Room: [['R101', 'غرفة 101', 'B01']],
  Department: [['ADMIN', 'الإدارة']], Employee: [['E001', 'أمين العهدة']],
  JournalName: [['FAJ', 'Fixed asset journal - يومية الأصول'], ['FADEP', 'Depreciation - الإهلاك']]
};

/** ينقل بيانات الإصدارات السابقة (شيت Lookups الموحّد) إلى الشيتات المستقلة ويملأ البيانات الافتراضية. */
function migrateV3_(ss) {
  const stamp = stamp_();
  const add = function (t, rec) { rec.CREATEDBY = 'setup'; rec.CREATEDON = stamp; appendRow_(t, rec); };
  const oldSheet = ss.getSheetByName('Lookups');
  const old = oldSheet ? readTable_('Lookups').rows.filter(function (r) { return r.CODE; }) : [];
  const legacy = function (type) { return old.filter(function (r) { return r.TYPE === type; }); };

  const at = readTable_('MainAccounts');
  if (!at.rows.some(function (r) { return r.MAINACCOUNTID; })) {
    if (legacy('MainAccount').length) {
      legacy('MainAccount').forEach(function (r) { add(at, { MAINACCOUNTID: r.CODE, NAME: r.NAME, ACCOUNTLEVEL: 'Posting', MAINACCOUNTTYPE: 'BalanceSheet' }); });
    } else {
      DEFAULT_ACCOUNTS.forEach(function (a) {
        add(at, { MAINACCOUNTID: a[0], NAME: a[1], ACCOUNTLEVEL: a[2], PARENTACCOUNT: a[3], MAINACCOUNTTYPE: a[4], MAINACCOUNTCATEGORY: a[5] });
      });
    }
  } else {
    at.rows.forEach(function (r, i) { if (r.MAINACCOUNTID && !r.ACCOUNTLEVEL) { r.ACCOUNTLEVEL = 'Posting'; writeRow_(at, i, r); } });
  }
  const accs = {};
  at.rows.forEach(function (r) { accs[low_(r.MAINACCOUNTID)] = 1; });
  const acc = function (c) { return c && accs[low_(c)] ? c : ''; };

  const gt = readTable_('AssetGroups');
  if (!gt.rows.some(function (r) { return r.FIXEDASSETGROUPID; })) {
    const src = legacy('AssetGroup').length ? legacy('AssetGroup').map(function (r) { return [r.CODE, r.NAME]; }) : DEFAULT_GROUPS;
    src.forEach(function (g) {
      const d = DEFAULT_GROUPS.filter(function (x) { return x[0] === g[0]; })[0] || [];
      add(gt, { FIXEDASSETGROUPID: g[0], NAME: g[1], ASSETACCOUNT: acc(d[2]), ACCUMDEPACCOUNT: acc(d[3]), DEPEXPENSEACCOUNT: acc(d[4]), DISPOSALACCOUNT: acc(d[5]) });
    });
  }

  LK.forEach(function (t) {
    const tt = readTable_(t.sheet);
    if (tt.rows.some(function (r) { return r.CODE; })) return;
    // عند الترقية تُنقل البيانات القديمة فقط؛ البيانات الافتراضية للتثبيت الجديد
    const src = legacy(t.v).length ? legacy(t.v).map(function (r) { return [r.CODE, r.NAME, r.PARENT || '']; }) : (oldSheet ? [] : (DEFAULT_LK[t.v] || []));
    src.forEach(function (x) {
      const rec = { CODE: x[0], NAME: x[1] || x[0] };
      if (t.parent) rec.PARENT = x[2] || '';
      add(tt, rec);
    });
  });
  if (oldSheet && !ss.getSheetByName('Lookups_archive')) oldSheet.setName('Lookups_archive');

  // الوحدات أصبحت على الأصل: تُنقل من الدفتر الرئيسي إن لم تكن مسجلة على الأصل
  const ast = readTable_('FixedAssets');
  const pb = primaryBooks_(readTable_('AssetBooks').rows);
  const groups = {};
  readTable_('AssetGroups').rows.forEach(function (g) { groups[low_(g.FIXEDASSETGROUPID)] = g; });
  ast.rows.forEach(function (a, i) {
    if (!a.FIXEDASSETNUMBER) return;
    let changed = false;
    const b = pb[low_(a.FIXEDASSETNUMBER)];
    if (!a.OPERATINGUNIT && !a.SUB2_OPERATINGUNIT && b && b.OPERATINGUNIT) {
      a.OPERATINGUNIT = b.OPERATINGUNIT; a.SUB1_OPERATINGUNIT = b.SUB1_OPERATINGUNIT || ''; a.SUB2_OPERATINGUNIT = b.SUB2_OPERATINGUNIT || '';
      changed = true;
    }
    const g = groups[low_(a.FIXEDASSETGROUPID)];
    if (!a.ASSETACCOUNT && g && g.ASSETACCOUNT) { a.ASSETACCOUNT = g.ASSETACCOUNT; changed = true; }
    if (changed) writeRow_(ast, i, a);
  });
}

/** الإصدار 6: اسم التطبيق الجديد + تصنيف حسابات النقدية للتدفقات النقدية. */
function migrateV6_() {
  const st = readTable_('Settings');
  st.rows.forEach(function (r, i) {
    if (r.KEY === 'ORGNAME' && r.VALUE === 'إدارة الأصول الثابتة') { r.VALUE = 'الإدارة المالية'; writeRow_(st, i, r); }
  });
  const at = readTable_('MainAccounts');
  at.rows.forEach(function (r, i) {
    if (r.MAINACCOUNTID && !r.CASHFLOWCATEGORY && r.ACCOUNTLEVEL !== 'Header' && String(r.MAINACCOUNTCATEGORY).toUpperCase() === 'CASH') {
      r.CASHFLOWCATEGORY = 'Cash'; writeRow_(at, i, r);
    }
  });
}

/** الإصدار 7: العملات، أسماء يوميات الإقفال، حساب الأرباح المحتجزة، تحويل سطور القيود القديمة للعملة المحلية. */
function migrateV7_() {
  const st = getSettings_();
  const local = String(st.CURRENCY || 'DJF').toUpperCase();
  const stamp = stamp_();
  const ct = readTable_('Currencies');
  const have = {};
  ct.rows.forEach(function (r) { have[String(r.CURRENCYCODE).toUpperCase()] = 1; });
  const names = { DJF: 'فرنك جيبوتي', USD: 'دولار أمريكي', EUR: 'يورو', SAR: 'ريال سعودي', AED: 'درهم إماراتي', EGP: 'جنيه مصري' };
  const fresh = !ct.rows.some(function (r) { return r.CURRENCYCODE; });
  [local].concat(fresh ? ['USD', 'EUR'] : []).forEach(function (c) {
    if (have[c]) return;
    have[c] = 1;
    appendRow_(ct, { CURRENCYCODE: c, NAME: names[c] || c, SYMBOL: '', CREATEDBY: 'setup', CREATEDON: stamp });
  });
  const jt = readTable_('JournalNames');
  const jn = {};
  jt.rows.forEach(function (r) { jn[low_(r.CODE)] = 1; });
  [['GJ', 'قيود عامة - General journal'], ['YE', 'إقفال السنة - Year-end closing']].forEach(function (x) {
    if (!jn[low_(x[0])]) appendRow_(jt, { CODE: x[0], NAME: x[1], CREATEDBY: 'setup', CREATEDON: stamp });
  });
  const at = readTable_('MainAccounts');
  const accs = {};
  at.rows.forEach(function (r) { accs[low_(r.MAINACCOUNTID)] = r; });
  if (accs['3100'] && !accs['310200']) {
    appendRow_(at, { MAINACCOUNTID: '310200', NAME: 'الأرباح المحتجزة', ACCOUNTLEVEL: 'Posting', PARENTACCOUNT: '3100', MAINACCOUNTTYPE: 'Equity',
      MAINACCOUNTCATEGORY: 'EQ', CREATEDBY: 'setup', CREATEDON: stamp });
    accs['310200'] = 1;
  }
  if (!st.RETAINEDEARNINGSACCOUNT && accs['310200']) setSettings_({ RETAINEDEARNINGSACCOUNT: '310200' });
  const gl = readTable_('GLJournalLines');
  gl.rows.forEach(function (l, i) {
    if (!l.LINEID || l.CURRENCYCODE) return;
    l.CURRENCYCODE = local; l.EXCHANGERATE = 1; l.AMOUNTCURDEBIT = num_(l.DEBITAMOUNT); l.AMOUNTCURCREDIT = num_(l.CREDITAMOUNT);
    writeRow_(gl, i, l);
  });
}

/** الإصدار 8: مجموعات البنوك، حسابات العملاء والأرصدة الافتتاحية، صندوق/بنك افتراضي، نوع الحساب لسطور القيود القديمة. */
function migrateV8_() {
  const stamp = stamp_();
  const local = String(getSettings_().CURRENCY || 'DJF').toUpperCase();
  const gt = readTable_('BankGroups');
  if (!gt.rows.some(function (r) { return r.CODE; })) {
    [['CASH', 'الصناديق'], ['BANK', 'البنوك']].forEach(function (x) { appendRow_(gt, { CODE: x[0], NAME: x[1], CREATEDBY: 'setup', CREATEDON: stamp }); });
  }
  const at = readTable_('MainAccounts');
  const accs = {};
  at.rows.forEach(function (r) { accs[low_(r.MAINACCOUNTID)] = r; });
  const add = function (id, name, lvl, parent, type, cat, cf) {
    if (accs[low_(id)]) return;
    appendRow_(at, { MAINACCOUNTID: id, NAME: name, ACCOUNTLEVEL: lvl, PARENTACCOUNT: parent, MAINACCOUNTTYPE: type, MAINACCOUNTCATEGORY: cat,
      CASHFLOWCATEGORY: cf || '', CREATEDBY: 'setup', CREATEDON: stamp });
    accs[low_(id)] = 1;
  };
  if (accs['1100']) {   // الدليل الافتراضي موجود: نضيف حسابات العملاء والأرصدة الافتتاحية
    add('1300', 'العملاء', 'Header', '', 'Asset', 'AR');
    add('130100', 'ذمم العملاء', 'Posting', '1300', 'Asset', 'AR');
    add('110200', 'الصندوق', 'Posting', '1100', 'Asset', 'CASH', 'Cash');
  }
  if (accs['3100']) add('310300', 'حساب الأرصدة الافتتاحية', 'Posting', '3100', 'Equity', 'EQ');
  if (!getSettings_().OPENINGBALANCEACCOUNT && accs['310300']) setSettings_({ OPENINGBALANCEACCOUNT: '310300' });
  const bt = readTable_('BankAccounts');
  if (!bt.rows.some(function (r) { return r.BANKACCOUNTID; })) {
    if (accs['110100']) appendRow_(bt, { BANKACCOUNTID: 'BANK01', NAME: 'البنك الرئيسي', BANKGROUP: 'BANK', CURRENCYCODE: local, LEDGERACCOUNT: '110100', CREATEDBY: 'setup', CREATEDON: stamp });
    if (accs['110200']) appendRow_(bt, { BANKACCOUNTID: 'CASH01', NAME: 'الصندوق الرئيسي', BANKGROUP: 'CASH', CURRENCYCODE: local, LEDGERACCOUNT: '110200', CREATEDBY: 'setup', CREATEDON: stamp });
  }
  const gl = readTable_('GLJournalLines');
  gl.rows.forEach(function (l, i) { if (l.LINEID && !l.ACCOUNTTYPE) { l.ACCOUNTTYPE = 'Ledger'; writeRow_(gl, i, l); } });
}

/** الإصدار 9: يوميات السندات، وتجميع سطور الموازنة القديمة (بدون رأس) في موازنات معتمدة. */
function migrateV9_() {
  const stamp = stamp_();
  const jt = readTable_('JournalNames');
  const have = {};
  jt.rows.forEach(function (r) { have[low_(r.CODE)] = 1; });
  [['RV', 'سندات القبض - Receipts'], ['PV', 'سندات الصرف - Payments']].forEach(function (x) {
    if (!have[low_(x[0])]) appendRow_(jt, { CODE: x[0], NAME: x[1], CREATEDBY: 'setup', CREATEDON: stamp });
  });
  const types = {};
  readTable_('MainAccounts').rows.forEach(function (a) { types[low_(a.MAINACCOUNTID)] = a.MAINACCOUNTTYPE; });
  const bt = readTable_('Budgets');
  const ht = readTable_('BudgetHeaders');
  const groups = {}, order = [];
  bt.rows.forEach(function (r, i) {
    if (!r.LINEID || r.BUDGETID) return;
    const type = types[low_(r.MAINACCOUNTID)] === 'Revenue' ? 'Revenue' : 'Expense';
    const k = r.BUDGETYEAR + '|' + (r.PERIODTYPE || 'Annual') + '|' + type;
    if (!groups[k]) { groups[k] = []; order.push(k); }
    groups[k].push(i);
  });
  order.forEach(function (k) {
    const p = k.split('|');
    const id = nextSeq_('BUDGETSEQ', 'BUDGETPREFIX', 'BG-');
    let total = 0;
    groups[k].forEach(function (i, n) {
      const r = bt.rows[i];
      r.BUDGETID = id; r.LINENUM = n + 1; r.BUDGETTYPE = p[2]; r.PERIODTYPE = p[1];
      total += num_(r.AMOUNT);
      writeRow_(bt, i, r);
    });
    appendRow_(ht, { BUDGETID: id, BUDGETYEAR: p[0], BUDGETTYPE: p[2], ENTRYDATE: today_(), PERIODTYPE: p[1], DESCRIPTION: 'منقولة من الإصدار 8',
      TOTALAMOUNT: round_(total, 3), STATUS: 'Approved', APPROVEDBY: 'setup', APPROVEDON: stamp, CREATEDBY: 'setup', CREATEDON: stamp });
  });
}

/** الإصدار 10: يوميات الفواتير، وحسابات المبيعات والمشتريات والضريبة في الدليل الافتراضي. */
function migrateV10_() {
  const stamp = stamp_();
  const jt = readTable_('JournalNames');
  const have = {};
  jt.rows.forEach(function (r) { have[low_(r.CODE)] = 1; });
  [['SI', 'فواتير المبيعات - Sales invoices'], ['PI', 'فواتير المشتريات - Purchase invoices']].forEach(function (x) {
    if (!have[low_(x[0])]) appendRow_(jt, { CODE: x[0], NAME: x[1], CREATEDBY: 'setup', CREATEDON: stamp });
  });
  const at = readTable_('MainAccounts');
  const accs = {};
  at.rows.forEach(function (r) { accs[low_(r.MAINACCOUNTID)] = 1; });
  const add = function (id, name, lvl, parent, type, cat) {
    if (accs[low_(id)]) return;
    appendRow_(at, { MAINACCOUNTID: id, NAME: name, ACCOUNTLEVEL: lvl, PARENTACCOUNT: parent, MAINACCOUNTTYPE: type, MAINACCOUNTCATEGORY: cat, CREATEDBY: 'setup', CREATEDON: stamp });
    accs[low_(id)] = 1;
  };
  if (!accs['4100'] || !accs['1100']) return;   // دليل حسابات مخصّص: لا نضيف شيئاً
  add('1400', 'أرصدة مدينة أخرى', 'Header', '', 'Asset', 'OTH');
  add('140100', 'ضريبة القيمة المضافة - مدخلات', 'Posting', '1400', 'Asset', 'TAX');
  add('2200', 'الضرائب المستحقة', 'Header', '', 'Liability', 'TAX');
  add('220100', 'ضريبة القيمة المضافة - مخرجات', 'Posting', '2200', 'Liability', 'TAX');
  add('410200', 'إيرادات المبيعات', 'Posting', '4100', 'Revenue', 'REV');
  add('6200', 'المصروفات التشغيلية', 'Header', '', 'Expense', 'OPEX');
  add('620100', 'مشتريات ومصروفات عامة', 'Posting', '6200', 'Expense', 'OPEX');
  const st = getSettings_(), o = {};
  if (!st.SALESACCOUNT) o.SALESACCOUNT = '410200';
  if (!st.PURCHASEACCOUNT) o.PURCHASEACCOUNT = '620100';
  if (!st.SALESTAXACCOUNT) o.SALESTAXACCOUNT = '220100';
  if (!st.PURCHTAXACCOUNT) o.PURCHTAXACCOUNT = '140100';
  if (Object.keys(o).length) setSettings_(o);
}

/** الإصدار 11: وحدات الأصناف، مجموعة أصناف افتراضية، الموازنات المعتمدة سابقاً تصبح مُرحّلة، ووحدة الفاتورة في الرأس. */
function migrateV11_() {
  const stamp = stamp_();
  const ut = readTable_('ItemUnits');
  if (!ut.rows.some(function (r) { return r.CODE; })) {
    [['PCS', 'قطعة'], ['BOX', 'علبة'], ['KG', 'كيلوجرام'], ['L', 'لتر'], ['SRV', 'خدمة'], ['HR', 'ساعة']].forEach(function (x) {
      appendRow_(ut, { CODE: x[0], NAME: x[1], CREATEDBY: 'setup', CREATEDON: stamp });
    });
  }
  const gt = readTable_('ItemGroups');
  const st = getSettings_();
  if (!gt.rows.some(function (r) { return r.ITEMGROUPID; }) && (st.SALESACCOUNT || st.PURCHASEACCOUNT)) {
    appendRow_(gt, { ITEMGROUPID: 'GEN', NAME: 'أصناف عامة', SALESACCOUNT: st.SALESACCOUNT || '', PURCHASEACCOUNT: st.PURCHASEACCOUNT || '', CREATEDBY: 'setup', CREATEDON: stamp });
  }
  // الموازنات المعتمدة قبل هذا الإصدار كانت تدخل المراقبة مباشرة — تُرحَّل حتى لا يتغير التقرير
  const bh = readTable_('BudgetHeaders');
  bh.rows.forEach(function (h, i) {
    if (h.BUDGETID && h.STATUS === 'Approved' && !h.POSTEDON) { h.STATUS = 'Posted'; h.POSTEDBY = 'setup'; h.POSTEDON = stamp; writeRow_(bh, i, h); }
  });
  // الفواتير القديمة: الوحدة من أول سطر إلى الرأس
  const it = readTable_('Invoices'), lt = readTable_('InvoiceLines');
  const first = {};
  lt.rows.forEach(function (l) { if (l.INVOICEID && l.SUB2_OPERATINGUNIT && !first[low_(l.INVOICEID)]) first[low_(l.INVOICEID)] = l; });
  it.rows.forEach(function (h, i) {
    const l = h.INVOICEID && !h.SUB2_OPERATINGUNIT && first[low_(h.INVOICEID)];
    if (l) { h.SUB2_OPERATINGUNIT = l.SUB2_OPERATINGUNIT; h.SUB1_OPERATINGUNIT = l.SUB1_OPERATINGUNIT; h.OPERATINGUNIT = l.OPERATINGUNIT; writeRow_(it, i, h); }
  });
}

/** الإصدار 12: يوميتا الإشعارات الدائنة والمدينة. */
function migrateV12_() {
  const jt = readTable_('JournalNames');
  const have = {};
  jt.rows.forEach(function (r) { have[low_(r.CODE)] = 1; });
  [['CN', 'إشعارات دائنة - Credit notes'], ['DN', 'إشعارات مدينة - Debit notes']].forEach(function (x) {
    if (!have[low_(x[0])]) appendRow_(jt, { CODE: x[0], NAME: x[1], CREATEDBY: 'setup', CREATEDON: stamp_() });
  });
}

/** الإصدار 13: حسابا مردودات المبيعات والمشتريات (في الدليل الافتراضي) وربطهما بالإعدادات والمجموعة العامة. */
function migrateV13_() {
  const stamp = stamp_();
  const at = readTable_('MainAccounts');
  const accs = {};
  at.rows.forEach(function (r) { accs[low_(r.MAINACCOUNTID)] = 1; });
  const add = function (id, name, parent, type) {
    if (accs[low_(id)] || !accs[low_(parent)]) return;
    appendRow_(at, { MAINACCOUNTID: id, NAME: name, ACCOUNTLEVEL: 'Posting', PARENTACCOUNT: parent, MAINACCOUNTTYPE: type, MAINACCOUNTCATEGORY: 'RET', CREATEDBY: 'setup', CREATEDON: stamp });
    accs[low_(id)] = 1;
  };
  add('410900', 'مردودات ومسموحات المبيعات', '4100', 'Revenue');
  add('620900', 'مردودات ومسموحات المشتريات', '6200', 'Expense');
  const st = getSettings_(), o = {};
  if (!st.SALESRETURNACCOUNT && accs['410900']) o.SALESRETURNACCOUNT = '410900';
  if (!st.PURCHRETURNACCOUNT && accs['620900']) o.PURCHRETURNACCOUNT = '620900';
  if (Object.keys(o).length) setSettings_(o);
  const gt = readTable_('ItemGroups');
  gt.rows.forEach(function (g, i) {
    if (low_(g.ITEMGROUPID) !== 'gen' || g.SALESRETURNACCOUNT || g.PURCHRETURNACCOUNT) return;
    g.SALESRETURNACCOUNT = o.SALESRETURNACCOUNT || st.SALESRETURNACCOUNT || '';
    g.PURCHRETURNACCOUNT = o.PURCHRETURNACCOUNT || st.PURCHRETURNACCOUNT || '';
    writeRow_(gt, i, g);
  });
}

/** بيانات تجريبية اختيارية مطابقة للصور (BLD_00001). */
function seedDemoData() {
  ensureSetup_();
  withLock_(function () {
    const at = readTable_('FixedAssets');
    if (!at.rows.some(function (r) { return r.FIXEDASSETNUMBER === 'BLD_00001'; })) {
      appendRow_(at, { FIXEDASSETNUMBER: 'BLD_00001', FIXEDASSETGROUPID: 'BLD', NAME: 'غرفة الومتال +تشطيب السقف', ASSETTYPE: 'Tangible',
        QUANTITY: 1, MODEL: '', GUARANTEEDATE: '', LOCATION: 'MAIN', BUILDING: 'B01', ROOMNUMBER: 'R101', DEPARTMENT: 'ADMIN', RESPONSIBLE: 'E001',
        ASSETACCOUNT: '120100', OPERATINGUNIT: 'HE01', SUB1_OPERATINGUNIT: 'DJI_HE0101', SUB2_OPERATINGUNIT: 'DJI_HE010001',
        CREATEDBY: 'setup', CREATEDON: stamp_() });
    }
    const bt = readTable_('AssetBooks');
    if (!bt.rows.some(function (r) { return r.FIXEDASSETNUMBER === 'BLD_00001'; })) {
      appendRow_(bt, { FIXEDASSETNUMBER: 'BLD_00001', BOOKID: 'General', ACQUISITIONDATE: '2020-09-01', ACQUISITIONPRICE: 730000,
        STATUS: 'Open', SERVICELIFE: 50, DEPRECIATIONPERIODS: 600, DEPRECIATIONPERIODSREMAINING: 465, SCRAPVALUE: 1,
        ACCUMULATEDDEPRECIATION: 164250, NETBOOKVALUE: 565750,
        DEPRECIATIONSTARTDATE: '2020-09-01', LASTDEPRECIATIONDATE: '2026-08-31', OPERATINGUNIT: 'HE01',
        SUB1_OPERATINGUNIT: 'DJI_HE0101', SUB2_OPERATINGUNIT: 'DJI_HE010001',
        DEFAULTDIMENSIONDISPLAYVALUE: 'HE01-DJI_HE0101-DJI_HE010001', CREATEDBY: 'setup', CREATEDON: stamp_() });
    }
  });
}

function ensureSheet_(ss, name, cols, numberCols) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const lastCol = sh.getLastColumn();
  const head = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  const missing = cols.filter(function (c) { return head.indexOf(c) < 0; });
  if (missing.length) {
    sh.getRange(1, head.length + 1, 1, missing.length).setValues([missing]);
  }
  const all = head.concat(missing);
  const rows = sh.getMaxRows();
  all.forEach(function (c, i) {
    sh.getRange(1, i + 1, rows, 1).setNumberFormat(numberCols.indexOf(c) >= 0 ? '#,##0.###' : '@');
  });
  sh.getRange(1, 1, 1, all.length).setFontWeight('bold').setBackground('#1f2a44').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

function cols_(entity) { return SCHEMAS[entity].fields.map(function (f) { return f.k; }).concat(AUDIT); }

/* ============================== Sheet helpers ============================== */

function sheetByName_(name) {
  const sh = getSs_().getSheetByName(name);
  if (!sh) throw new Error('الشيت غير موجود: ' + name + ' — شغّل الدالة setup');
  return sh;
}

function readTable_(name) {
  const sh = sheetByName_(name);
  const vals = sh.getDataRange().getValues();
  const head = (vals.shift() || []).map(String);
  const tz = Session.getScriptTimeZone();
  const rows = vals.map(function (r) {
    const obj = {};
    head.forEach(function (h, i) {
      let v = r[i];
      if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
      obj[h] = v;
    });
    return obj;
  });
  return { sh: sh, head: head, rows: rows };
}

function cellVal_(v) {
  if (v === undefined || v === null) return '';
  if (typeof v === 'string' && /^[=+@]/.test(v)) return "'" + v; // منع حقن الصيغ
  return v;
}

function toRow_(head, rec) { return head.map(function (h) { return cellVal_(rec[h]); }); }

function writeRow_(t, idx, rec) {
  t.sh.getRange(idx + 2, 1, 1, t.head.length).setValues([toRow_(t.head, rec)]);
  t.rows[idx] = rec;
}

function appendRow_(t, rec) {
  const sh = t.sh;
  const r = t.rows.length + 2;
  if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 200);
  sh.getRange(r, 1, 1, t.head.length).setValues([toRow_(t.head, rec)]);
  t.rows.push(rec);
  return t.rows.length - 1;
}

function deleteRows_(t, idxs) {
  idxs.sort(function (a, b) { return b - a; }).forEach(function (i) {
    t.sh.deleteRow(i + 2);
    t.rows.splice(i, 1);
  });
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function stamp_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'); }
function today_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function low_(s) { return String(s === undefined || s === null ? '' : s).trim().toLowerCase(); }
function norm_(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9؀-ۿ]/g, ''); }
function num_(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
function round_(n, d) { const p = Math.pow(10, d || 3); return Math.round(n * p) / p; }
function isTrue_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function blank_(r) { return !Object.keys(r).some(function (k) { return r[k] !== '' && r[k] !== null; }); }

/* ============================== Settings ============================== */

function getSettings_() {
  const cache = CacheService.getScriptCache();
  const c = cache.get('SETTINGS');
  if (c) return JSON.parse(c);
  const out = Object.assign({}, SETTING_DEFAULTS);
  readTable_('Settings').rows.forEach(function (r) { if (r.KEY) out[r.KEY] = String(r.VALUE); });
  const json = JSON.stringify(out);
  if (json.length < 95000) cache.put('SETTINGS', json, 600);
  return out;
}

function setSettings_(obj) {
  const t = readTable_('Settings');
  Object.keys(obj).forEach(function (k) {
    const i = t.rows.findIndex(function (r) { return r.KEY === k; });
    if (i >= 0) writeRow_(t, i, { KEY: k, VALUE: String(obj[k]) });
    else appendRow_(t, { KEY: k, VALUE: String(obj[k]) });
  });
  CacheService.getScriptCache().remove('SETTINGS');
}

function publicSettings_() {
  const s = getSettings_();
  const out = {};
  PUBLIC_SETTINGS.forEach(function (k) { out[k] = s[k]; });
  return out;
}

function getPublicInfo() {
  ensureSetup_();
  const s = getSettings_();
  return { orgName: s.ORGNAME, logo: s.LOGO, version: APP_VERSION };
}

/* ============================== Security ============================== */

function hash_(pw, salt) {
  let h = salt + '|' + pw;
  let b64 = '';
  for (let i = 0; i < HASH_ROUNDS; i++) {
    b64 = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8));
    h = b64 + salt;
  }
  return 'sha256$' + HASH_ROUNDS + '$' + b64;
}

function sha_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

function safeEq_(a, b) {
  a = String(a); b = String(b);
  let d = a.length ^ b.length;
  for (let i = 0; i < Math.min(a.length, b.length); i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}

function checkPassword_(pw) {
  pw = String(pw || '');
  if (pw.length < 8 || !/[A-Za-z؀-ۿ]/.test(pw) || !/\d/.test(pw)) {
    throw new Error('كلمة المرور يجب أن تكون 8 أحرف على الأقل وتحتوي على حروف وأرقام');
  }
  return pw;
}

function roleTemplate_(role) {
  const p = {};
  PERM_PAGES.forEach(function (pg) { p[pg] = {}; PERM_ACTIONS.forEach(function (a) { p[pg][a] = false; }); });
  const set = function (pages, acts) { pages.forEach(function (pg) { acts.forEach(function (a) { p[pg][a] = true; }); }); };
  const data = ['assets', 'books', 'journal', 'vendors', 'masterdata', 'customers', 'banks', 'budget', 'cashvouchers', 'invoices', 'items'];
  if (role === 'SuperAdmin' || role === 'Admin') set(PERM_PAGES, PERM_ACTIONS);
  else if (role === 'Accountant') { set(data.concat(['transfers', 'gl']), ['view', 'add', 'edit']); set(['fiscal'], ['view']); set(['depreciation'], ['view', 'add']); set(['reports'], ['view']); set(['cashvouchers', 'invoices', 'budget'], ['post']); }
  else if (role !== 'none') set(data.concat(['reports', 'transfers', 'gl']), ['view']);
  return p;
}

function effectivePerms_(user) {
  if (user.ROLE === 'SuperAdmin') return roleTemplate_('SuperAdmin');
  const base = roleTemplate_('none');
  let saved = null;
  try { saved = user.PERMISSIONS ? JSON.parse(user.PERMISSIONS) : null; } catch (e) { saved = null; }
  if (!saved) return roleTemplate_(user.ROLE);
  const tpl = roleTemplate_(user.ROLE);
  PERM_PAGES.forEach(function (pg) {
    // صفحة جديدة لم تكن موجودة عند حفظ صلاحيات المستخدم: تؤخذ من قالب الدور
    PERM_ACTIONS.forEach(function (a) { base[pg][a] = saved[pg] ? !!saved[pg][a] : tpl[pg][a]; });
  });
  return base;
}

function can_(s, page, action) { return !!(s.perms[page] && s.perms[page][action]); }

function need_(s, page, action) {
  if (!can_(s, page, action)) {
    const names = { view: 'الدخول', add: 'الإضافة', edit: 'التعديل', delete: 'الحذف', approve: 'الاعتماد', post: 'الترحيل' };
    throw new Error('ليس لديك صلاحية ' + (names[action] || action) + ' في هذه الصفحة');
  }
}

/* ============================== Sessions ============================== */

function findUser_(username) {
  const t = readTable_('Users');
  const i = t.rows.findIndex(function (r) { return low_(r.USERNAME) === low_(username); });
  return i < 0 ? null : { t: t, idx: i, u: t.rows[i] };
}

function getUserCached_(username) {
  const cache = CacheService.getScriptCache();
  const c = cache.get('U_' + username);
  if (c) return JSON.parse(c);
  const f = findUser_(username);
  if (!f) return null;
  cache.put('U_' + username, JSON.stringify(f.u), 60);
  return f.u;
}

function clearUserCache_(username) { CacheService.getScriptCache().remove('U_' + low_(username)); }

function sessionCacheSecs_() {
  return Math.min(21600, Number(getSettings_().SESSIONTIMEOUT || 60) * 60 + 60);
}

function createSession_(username) {
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const sid = sha_(token);
  const now = new Date().toISOString();
  const max = Math.max(1, Number(getSettings_().MAXSESSIONS || 3));
  withLock_(function () {
    const t = readTable_('Sessions');
    const mine = [];
    t.rows.forEach(function (r, i) { if (r.USERNAME === username && isTrue_(r.ACTIVE)) mine.push(i); });
    mine.sort(function (a, b) { return String(t.rows[a].LASTSEEN).localeCompare(String(t.rows[b].LASTSEEN)); });
    while (mine.length >= max) {
      const i = mine.shift();
      const r = t.rows[i];
      r.ACTIVE = 'FALSE';
      writeRow_(t, i, r);
      CacheService.getScriptCache().remove('S_' + r.SESSIONID);
    }
    appendRow_(t, { SESSIONID: sid, USERNAME: username, CREATEDON: now, LASTSEEN: now, ACTIVE: 'TRUE', CLIENT: '' });
  });
  CacheService.getScriptCache().put('S_' + sid, JSON.stringify({ u: username, ls: Date.now(), w: Date.now() }), sessionCacheSecs_());
  return token;
}

function endSession_(sid) {
  CacheService.getScriptCache().remove('S_' + sid);
  const t = readTable_('Sessions');
  const i = t.rows.findIndex(function (r) { return r.SESSIONID === sid; });
  if (i >= 0 && isTrue_(t.rows[i].ACTIVE)) { t.rows[i].ACTIVE = 'FALSE'; writeRow_(t, i, t.rows[i]); }
}

function endUserSessions_(username) {
  const t = readTable_('Sessions');
  t.rows.forEach(function (r, i) {
    if (low_(r.USERNAME) === low_(username) && isTrue_(r.ACTIVE)) {
      r.ACTIVE = 'FALSE';
      writeRow_(t, i, r);
      CacheService.getScriptCache().remove('S_' + r.SESSIONID);
    }
  });
}

function auth_(token) {
  if (!token) throw new Error(AUTH_ERR + 'يرجى تسجيل الدخول');
  const sid = sha_(token);
  const cache = CacheService.getScriptCache();
  const timeoutMs = Number(getSettings_().SESSIONTIMEOUT || 60) * 60000;
  const now = Date.now();
  let s = JSON.parse(cache.get('S_' + sid) || 'null');
  if (!s) {
    const t = readTable_('Sessions');
    const r = t.rows.find(function (x) { return x.SESSIONID === sid; });
    if (!r || !isTrue_(r.ACTIVE)) throw new Error(AUTH_ERR + 'انتهت الجلسة، يرجى تسجيل الدخول مجدداً');
    const ls = Date.parse(r.LASTSEEN) || 0;
    s = { u: r.USERNAME, ls: ls, w: ls };
  }
  if (now - s.ls > timeoutMs) {
    endSession_(sid);
    throw new Error(AUTH_ERR + 'انتهت الجلسة بسبب عدم النشاط، يرجى تسجيل الدخول مجدداً');
  }
  s.ls = now;
  if (now - s.w > 5 * 60000) {
    const t = readTable_('Sessions');
    const i = t.rows.findIndex(function (x) { return x.SESSIONID === sid; });
    if (i < 0 || !isTrue_(t.rows[i].ACTIVE)) { cache.remove('S_' + sid); throw new Error(AUTH_ERR + 'تم إنهاء الجلسة'); }
    t.rows[i].LASTSEEN = new Date(now).toISOString();
    writeRow_(t, i, t.rows[i]);
    s.w = now;
  }
  cache.put('S_' + sid, JSON.stringify(s), sessionCacheSecs_());
  const user = getUserCached_(s.u);
  if (!user || user.STATUS !== 'Active') {
    endSession_(sid);
    throw new Error(AUTH_ERR + 'الحساب غير نشط');
  }
  return { sid: sid, username: user.USERNAME, user: user, perms: effectivePerms_(user) };
}

/* ============================== Auth API ============================== */

function login(username, password) {
  ensureSetup_();
  username = low_(username);
  const cache = CacheService.getScriptCache();
  const fk = 'F_' + username;
  const fails = Number(cache.get(fk) || 0);
  if (fails >= 5) throw new Error('تم إيقاف الدخول مؤقتاً بعد 5 محاولات فاشلة. حاول بعد 15 دقيقة.');
  const f = findUser_(username);
  const ok = f && f.u.SALT && safeEq_(hash_(String(password || ''), f.u.SALT), f.u.PASSWORDHASH);
  if (!ok || f.u.STATUS !== 'Active') {
    cache.put(fk, String(fails + 1), 900);
    throw new Error(ok ? 'الحساب غير نشط، تواصل مع مدير النظام' : 'اسم المستخدم أو كلمة المرور غير صحيحة');
  }
  cache.remove(fk);
  const token = createSession_(f.u.USERNAME);
  withLock_(function () {
    const g = findUser_(username);
    g.u.LASTLOGIN = stamp_();
    writeRow_(g.t, g.idx, g.u);
  });
  clearUserCache_(username);
  return bootstrap(token);
}

function logout(token) {
  if (token) endSession_(sha_(token));
  return true;
}

function bootstrap(token) {
  const s = auth_(token);
  const clientSchemas = {};
  Object.keys(SCHEMAS).forEach(function (e) {
    const x = SCHEMAS[e];
    clientSchemas[e] = { perm: x.perm, key: x.key, title: x.title, sections: x.sections, fields: x.fields, autoId: x.autoId || '',
      lkType: x.lkType || '', parentType: x.parentType || '' };
  });
  return {
    token: token,
    version: APP_VERSION,
    user: { username: s.user.USERNAME, fullName: s.user.FULLNAME, email: s.user.EMAIL, role: s.user.ROLE, mustChange: isTrue_(s.user.MUSTCHANGE) },
    perms: s.perms,
    settings: publicSettings_(),
    schemas: clientSchemas,
    lookups: lookups_(),
    refs: refs_(s),
    roles: ROLES,
    lookupTypes: LK.map(function (t) { return { v: t.v, ar: t.ar, en: t.en, parent: t.parent || '' }; }),
    permPages: PERM_PAGES,
    permActions: PERM_ACTIONS
  };
}

function changePassword(token, oldPw, newPw) {
  const s = auth_(token);
  checkPassword_(newPw);
  return withLock_(function () {
    const f = findUser_(s.username);
    if (!safeEq_(hash_(String(oldPw || ''), f.u.SALT), f.u.PASSWORDHASH)) throw new Error('كلمة المرور الحالية غير صحيحة');
    if (oldPw === newPw) throw new Error('كلمة المرور الجديدة يجب أن تختلف عن الحالية');
    const salt = Utilities.getUuid();
    f.u.SALT = salt; f.u.PASSWORDHASH = hash_(newPw, salt); f.u.MUSTCHANGE = 'FALSE';
    f.u.MODIFIEDBY = s.username; f.u.MODIFIEDON = stamp_();
    writeRow_(f.t, f.idx, f.u);
    clearUserCache_(s.username);
    return true;
  });
}

/* ============================== Reference data ============================== */

function lookups_() {
  const out = {};
  const rd = function (sheet, code, name, parent) {
    return readTable_(sheet).rows.filter(function (r) { return r[code]; })
      .map(function (r) { return { c: String(r[code]), n: String(r[name] || ''), p: String(parent ? r[parent] || '' : '') }; })
      .sort(function (a, b) { return a.c.localeCompare(b.c); });
  };
  LK.forEach(function (t) { out[t.v] = rd(t.sheet, 'CODE', 'NAME', t.parent ? 'PARENT' : ''); });
  out.AssetGroup = rd('AssetGroups', 'FIXEDASSETGROUPID', 'NAME', '');
  return out;
}

function refs_(s) {
  const out = { assets: [], vendors: [], books: [], accounts: [], groups: [], currencies: [], rates: [], customers: [], banks: [] };
  if (['customers', 'gl', 'reports', 'cashvouchers', 'invoices'].some(function (pg) { return can_(s, pg, 'view'); })) {
    out.customers = readTable_('Customers').rows.filter(function (r) { return r.CUSTOMERACCOUNT; })
      .map(function (r) { return { c: String(r.CUSTOMERACCOUNT), n: String(r.NAME || ''), acc: r.LEDGERACCOUNT || '', cur: r.CURRENCYCODE || '',
        ad: String(r.ADDRESS || ''), ph: String(r.PHONE || ''), em: String(r.EMAIL || '') }; });
  }
  if (['banks', 'gl', 'reports', 'cashvouchers'].some(function (pg) { return can_(s, pg, 'view'); })) {
    out.banks = readTable_('BankAccounts').rows.filter(function (r) { return r.BANKACCOUNTID; })
      .map(function (r) { return { c: String(r.BANKACCOUNTID), n: String(r.NAME || ''), acc: r.LEDGERACCOUNT || '', cur: r.CURRENCYCODE || '' }; });
  }
  if (['invoices', 'cashvouchers'].some(function (pg) { return can_(s, pg, 'view'); })) {
    out.invoices = readTable_('Invoices').rows.filter(function (r) {
      return r.INVOICEID && r.STATUS === 'Posted' && INV_KIND[r.INVOICETYPE] && !INV_KIND[r.INVOICETYPE].credit && num_(r.TOTALAMOUNT) - num_(r.CREDITAMOUNT) > 0.0005;
    }).map(function (r) { return { c: String(r.INVOICEID), n: String(r.INVOICEDATE) + ' — ' + Math.max(0, invRem_(r)) + ' ' + r.CURRENCYCODE,
        t: INV_KIND[r.INVOICETYPE].side, k: r.INVOICETYPE, p: String(r.PARTYID), rem: invRem_(r), cap: round_(num_(r.TOTALAMOUNT) - num_(r.CREDITAMOUNT), 3),
        cur: r.CURRENCYCODE, rate: num_(r.EXCHANGERATE) || 1 }; });
  }
  if (['items', 'invoices'].some(function (pg) { return can_(s, pg, 'view'); })) {
    out.itemgroups = readTable_('ItemGroups').rows.filter(function (r) { return r.ITEMGROUPID; })
      .map(function (r) { return { c: String(r.ITEMGROUPID), n: String(r.NAME || ''), sa: r.SALESACCOUNT || '', pa: r.PURCHASEACCOUNT || '',
        sr: r.SALESRETURNACCOUNT || '', pr: r.PURCHRETURNACCOUNT || '' }; });
    ['salesitems', 'purchitems'].forEach(function (e) {
      out[e] = readTable_(SCHEMAS[e].sheet).rows.filter(function (r) { return r.ITEMID; })
        .map(function (r) { return { c: String(r.ITEMID), n: String(r.NAME || ''), g: String(r.ITEMGROUP || ''), u: String(r.UNIT || ''), pr: num_(r.PRICE) }; });
    });
  }
  out.currencies = readTable_('Currencies').rows.filter(function (r) { return r.CURRENCYCODE; })
    .map(function (r) { return { c: String(r.CURRENCYCODE), n: String(r.NAME || '') }; });
  out.rates = readTable_('ExchangeRates').rows.filter(function (r) { return r.CURRENCYCODE; })
    .map(function (r) { return { c: String(r.CURRENCYCODE), d: String(r.FROMDATE), r: num_(r.RATE) }; });
  out.accounts = readTable_('MainAccounts').rows.filter(function (r) { return r.MAINACCOUNTID; })
    .map(function (r) { return { c: String(r.MAINACCOUNTID), n: String(r.NAME || ''), p: String(r.PARENTACCOUNT || ''), h: r.ACCOUNTLEVEL === 'Header' ? 1 : 0, t: r.MAINACCOUNTTYPE || '' }; });
  out.groups = readTable_('AssetGroups').rows.filter(function (r) { return r.FIXEDASSETGROUPID; }).map(function (r) {
    return { c: String(r.FIXEDASSETGROUPID), n: String(r.NAME || ''), a: r.ASSETACCOUNT || '', ad: r.ACCUMDEPACCOUNT || '',
      de: r.DEPEXPENSEACCOUNT || '', di: r.DISPOSALACCOUNT || '' };
  });
  if (['assets', 'books', 'journal', 'depreciation', 'reports'].some(function (pg) { return can_(s, pg, 'view'); })) {
    out.assets = readTable_('FixedAssets').rows.filter(function (r) { return r.FIXEDASSETNUMBER; })
      .map(function (r) {
        return { c: String(r.FIXEDASSETNUMBER), n: String(r.NAME || ''), g: r.FIXEDASSETGROUPID || '',
          ou: r.OPERATINGUNIT || '', s1: r.SUB1_OPERATINGUNIT || '', s2: r.SUB2_OPERATINGUNIT || '',
          loc: r.LOCATION || '', bld: r.BUILDING || '', room: r.ROOMNUMBER || '', dept: r.DEPARTMENT || '', emp: r.RESPONSIBLE || '' };
      });
    out.books = readTable_('AssetBooks').rows.filter(function (r) { return r.FIXEDASSETNUMBER; })
      .map(function (r) { return { a: String(r.FIXEDASSETNUMBER), c: String(r.BOOKID), n: String(r.STATUS || '') }; });
  }
  if (['vendors', 'journal', 'gl', 'cashvouchers', 'invoices'].some(function (pg) { return can_(s, pg, 'view'); })) {
    out.vendors = readTable_('Vendors').rows.filter(function (r) { return r.VENDORACCOUNTNUMBER; })
      .map(function (r) { return { c: String(r.VENDORACCOUNTNUMBER), n: String(r.VENDORORGANIZATIONNAME || ''), acc: r.VENDORLEDGERACCOUNT || '', cur: r.CURRENCYCODE || '',
        ad: String(r.ADDRESS || ''), em: String(r.PRIMARYEMAILADDRESS || '') }; });
  }
  return out;
}

function getRefs(token) {
  const s = auth_(token);
  return { lookups: lookups_(), refs: refs_(s) };
}

/* ============================== Data cleaning ============================== */

function toNum_(v, f) {
  if (v === '' || v === null || v === undefined) return '';
  if (typeof v === 'number') return v;
  let s = String(v).trim().replace(/[\s,]/g, '').replace(/[٠-٩]/g, function (d) { return String(d.charCodeAt(0) - 1632); });
  let neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
  if (s === '') return '';
  const n = Number(s);
  if (isNaN(n)) throw new Error('قيمة رقمية غير صحيحة في الحقل "' + f.ar + '": ' + v);
  return neg ? -n : n;
}

function toDate_(v, fmt, f) {
  if (v === '' || v === null || v === undefined) return '';
  const tz = Session.getScriptTimeZone();
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  if (typeof v === 'number') return Utilities.formatDate(new Date(Math.round((v - 25569) * 86400000)), 'UTC', 'yyyy-MM-dd');
  const s = String(v).trim();
  let m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/);
  if (m) {
    let mo = m[1], d = m[2];
    if (fmt === 'DMY') { mo = m[2]; d = m[1]; }
    if (Number(mo) > 12) { const x = mo; mo = d; d = x; }
    return m[3] + '-' + ('0' + mo).slice(-2) + '-' + ('0' + d).slice(-2);
  }
  const p = Date.parse(s);
  if (!isNaN(p)) return Utilities.formatDate(new Date(p), tz, 'yyyy-MM-dd');
  throw new Error('تاريخ غير صحيح في الحقل "' + f.ar + '": ' + v);
}

function toOpt_(v, f) {
  if (v === '' || v === null || v === undefined) return '';
  const n = norm_(v);
  const hit = f.opts.find(function (op) { return norm_(op.v) === n || norm_(op.en) === n || norm_(op.ar) === n; });
  if (!hit) throw new Error('قيمة غير مقبولة في الحقل "' + f.ar + '": ' + v);
  return hit.v;
}

function clean_(sch, raw, dateFmt) {
  const out = {};
  sch.fields.forEach(function (f) {
    let v = raw ? raw[f.k] : '';
    if (v === undefined || v === null) v = '';
    if (f.type === 'number') v = toNum_(v, f);
    else if (f.type === 'date') v = toDate_(v, dateFmt, f);
    else if (f.type === 'select') v = toOpt_(v, f);
    else {
      v = String(v).trim();
      if (v.length > 1000) v = v.slice(0, 1000);
      if (f.type === 'email' && v && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)) throw new Error('بريد إلكتروني غير صحيح: ' + v);
    }
    out[f.k] = v;
  });
  return out;
}

function keyOf_(sch, r) { return sch.key.map(function (k) { return low_(r[k]); }).join('|'); }

function dimSplit_(s) { return String(s || '').split('-').map(function (x) { return x.trim(); }).filter(String); }
function dimJoin_(a) { return a.filter(function (x) { return x !== '' && x !== null && x !== undefined; }).join('-'); }

/* ============================== Context & hooks ============================== */

function ctx_(s, entity) {
  const tables = {};
  const index = {};
  const settings = getSettings_();
  const ctx = {
    s: s,
    dateFmt: settings.DATEFORMAT,
    table: function (e) { if (!tables[e]) tables[e] = readTable_(SCHEMAS[e].sheet); return tables[e]; },
    index: function (e) {
      if (!index[e]) {
        const m = {};
        ctx.table(e).rows.forEach(function (r, i) { if (!blank_(r)) m[keyOf_(SCHEMAS[e], r)] = i; });
        index[e] = m;
      }
      return index[e];
    },
    /** بحث بالمفتاح لكيان مفتاحه حقل واحد. */
    find: function (e, code) {
      if (code === '' || code === undefined || code === null) return null;
      const i = ctx.index(e)[low_(code)];
      return i === undefined ? null : ctx.table(e).rows[i];
    },
    asset: function (n) { return ctx.find('assets', n); },
    resetIndex: function (e) { delete index[e]; },
    voucherDate: '',
    /** سعر الصرف الساري: آخر سعر بتاريخ لا يتجاوز تاريخ الحركة. */
    rate: function (cur, date) {
      let best = null;
      ctx.table('exrates').rows.forEach(function (r) {
        if (low_(r.CURRENCYCODE) !== low_(cur) || !(num_(r.RATE) > 0)) return;
        if (date && String(r.FROMDATE) > String(date)) return;
        if (!best || String(r.FROMDATE) > String(best.FROMDATE)) best = r;
      });
      return best ? num_(best.RATE) : 0;
    },
    dirty: {},
    touch: function (e, row) {
      const i = ctx.table(e).rows.indexOf(row);
      if (i >= 0) (ctx.dirty[e] = ctx.dirty[e] || {})[i] = 1;
    },
    flush: function () {
      Object.keys(ctx.dirty).forEach(function (e) {
        const t = ctx.table(e);
        Object.keys(ctx.dirty[e]).forEach(function (i) { writeRow_(t, Number(i), t.rows[i]); });
      });
      ctx.dirty = {};
    },
    batchStatus: function (b) {
      let st = '';
      ctx.table('journal').rows.forEach(function (r) {
        if (low_(r.JOURNALBATCHNUMBER) !== low_(b)) return;
        if (r.STATUS === 'Posted') st = 'Posted';
        else if (r.STATUS === 'Approved' && st !== 'Posted') st = 'Approved';
        else if (!st) st = 'Draft';
      });
      return st;
    },
    importBatch: '',
    batch: function () { if (!ctx.importBatch) ctx.importBatch = nextBatch_(); return ctx.importBatch; },
    trfNo: '',
    transferNo: function () { if (!ctx.trfNo) ctx.trfNo = nextSeq_('TRANSFERSEQ', 'TRANSFERPREFIX', 'TRF-'); return ctx.trfNo; },
    docStatus: function (e, field, value) {
      let st = '';
      ctx.table(e).rows.forEach(function (r) {
        if (low_(r[field]) !== low_(value)) return;
        if (r.STATUS === 'Posted') st = 'Posted';
        else if (r.STATUS === 'Approved' && st !== 'Posted') st = 'Approved';
        else if (!st) st = 'Draft';
      });
      return st;
    }
  };
  return ctx;
}

function nextSeq_(seqKey, prefixKey, dflt) {
  const st = getSettings_();
  const n = Number(st[seqKey] || 0) + 1;
  const o = {};
  o[seqKey] = n;
  setSettings_(o);
  return (st[prefixKey] || dflt) + ('000000' + n).slice(-6);
}

function nextBatch_() { return nextSeq_('JOURNALSEQ', 'JOURNALPREFIX', 'FAJ-'); }

/** يملأ المستويات الأعلى من الشجرة بدءاً من أدنى مستوى مُدخل (مثلاً: الفرعية 2 ← الفرعية 1 ← الرئيسية). */
function deriveChain_(ctx, rec, chain) {
  for (let i = 0; i < chain.length; i++) {
    const v = rec[chain[i][0]];
    if (v === '' || v === undefined) continue;
    let row = ctx.find(chain[i][1], v);
    if (!row) throw new Error(LKMAP[chain[i][1]].ar + ': الكود غير موجود "' + v + '" — أضفه أولاً من البيانات الأساسية');
    rec[chain[i][0]] = row.CODE;
    for (let j = i + 1; j < chain.length && row && row.PARENT; j++) {
      const parent = ctx.find(chain[j][1], row.PARENT);
      rec[chain[j][0]] = parent ? parent.CODE : row.PARENT;
      row = parent;
    }
    return;
  }
}

function checkLk_(ctx, rec, field, type) {
  if (!rec[field]) return;
  const row = ctx.find(type, rec[field]);
  if (!row) throw new Error(LKMAP[type].ar + ': الكود غير موجود "' + rec[field] + '"');
  rec[field] = row.CODE;
}

/** يتحقق أن الحساب موجود في الدليل وأنه حساب فرعي (يقبل الترحيل). */
function postingAccount_(ctx, code, label) {
  if (!code) return '';
  const a = ctx.find('accounts', code);
  if (!a) throw new Error(label + ': الحساب غير موجود في دليل الحسابات "' + code + '"');
  if (a.ACCOUNTLEVEL === 'Header') throw new Error(label + ': الحساب ' + code + ' رئيسي، اختر حساباً فرعياً');
  return a.MAINACCOUNTID;
}

const DEP_TYPES = { Depreciation: 1, DepreciationAdj: 1, ExtraordinaryDepreciation: 1 };
const DISPOSAL_TYPES = { DisposalSale: 1, DisposalScrap: 1 };

/** حساب الأصل في القيد حسب نوع المعاملة (من مجموعة الأصل). */
function ledgerAccountFor_(type, g) {
  if (!g) return '';
  return DEP_TYPES[type] ? (g.ACCUMDEPACCOUNT || '') : (g.ASSETACCOUNT || '');
}

/** الحساب المقابل الافتراضي حسب نوع المعاملة (من مجموعة الأصل). */
function defaultOffsetFor_(type, g) {
  if (!g) return '';
  if (DEP_TYPES[type]) return g.DEPEXPENSEACCOUNT || '';
  if (DISPOSAL_TYPES[type]) return g.DISPOSALACCOUNT || '';
  return '';
}

function copyUnits_(from, to, prefixTo) {
  prefixTo = prefixTo || '';
  to[prefixTo + 'OPERATINGUNIT'] = from.OPERATINGUNIT || '';
  to[prefixTo + 'SUB1_OPERATINGUNIT'] = from.SUB1_OPERATINGUNIT || '';
  to[prefixTo + 'SUB2_OPERATINGUNIT'] = from.SUB2_OPERATINGUNIT || '';
}

const HOOKS = {
  assets: function (rec, ex, ctx) {
    if (rec.QUANTITY !== '' && rec.QUANTITY < 0) throw new Error('الكمية لا يمكن أن تكون سالبة');
    const g = ctx.find('groups', rec.FIXEDASSETGROUPID);
    if (!g) throw new Error('مجموعة الأصل غير موجودة: ' + rec.FIXEDASSETGROUPID + ' — أضفها من صفحة مجموعات الأصول');
    rec.FIXEDASSETGROUPID = g.FIXEDASSETGROUPID;
    rec.ASSETACCOUNT = g.ASSETACCOUNT || '';
    deriveChain_(ctx, rec, UNIT_CHAIN);
    deriveChain_(ctx, rec, LOCATION_CHAIN);
    checkLk_(ctx, rec, 'DEPARTMENT', 'Department');
    checkLk_(ctx, rec, 'RESPONSIBLE', 'Employee');
  },

  books: function (rec, ex, ctx) {
    const a = ctx.asset(rec.FIXEDASSETNUMBER);
    if (!a) throw new Error('رقم الأصل غير موجود: ' + rec.FIXEDASSETNUMBER);
    rec.FIXEDASSETNUMBER = a.FIXEDASSETNUMBER;
    if (rec.SERVICELIFE !== '' && rec.DEPRECIATIONPERIODS === '') rec.DEPRECIATIONPERIODS = Math.round(rec.SERVICELIFE * 12);
    if (rec.DEPRECIATIONPERIODS !== '' && rec.SERVICELIFE === '') rec.SERVICELIFE = round_(rec.DEPRECIATIONPERIODS / 12, 2);
    if (rec.DEPRECIATIONPERIODSREMAINING === '' && rec.DEPRECIATIONPERIODS !== '') rec.DEPRECIATIONPERIODSREMAINING = rec.DEPRECIATIONPERIODS;
    if (rec.DEPRECIATIONPERIODSREMAINING !== '' && rec.DEPRECIATIONPERIODS !== '' && rec.DEPRECIATIONPERIODSREMAINING > rec.DEPRECIATIONPERIODS) {
      throw new Error('فترات الإهلاك المتبقية أكبر من إجمالي فترات الإهلاك');
    }
    if (rec.ACQUISITIONPRICE !== '' && rec.ACQUISITIONPRICE < 0) throw new Error('سعر الاقتناء لا يمكن أن يكون سالباً');
    if (ex && ex.OPENINGVOUCHER && ctx.find('glheaders', ex.OPENINGVOUCHER) &&
        (num_(ex.ACQUISITIONPRICE) !== num_(rec.ACQUISITIONPRICE) || num_(ex.ACCUMULATEDDEPRECIATION) !== num_(rec.ACCUMULATEDDEPRECIATION))) {
      throw new Error('رصيد هذا الدفتر مُقيّد في القيد الافتتاحي ' + ex.OPENINGVOUCHER + ' — عدّل القيمة عبر يومية الأصول (تسوية اقتناء / تسوية إهلاك) أو احذف القيد وهو مسودة');
    }
    // الأبعاد تأتي من الأصل. إن لم تكن على الأصل (استيراد من D365) تُنقل من الدفتر إلى الأصل.
    if (a.OPERATINGUNIT || a.SUB2_OPERATINGUNIT) {
      copyUnits_(a, rec);
    } else {
      if (!rec.OPERATINGUNIT && rec.DEFAULTDIMENSIONDISPLAYVALUE) {
        const p = dimSplit_(rec.DEFAULTDIMENSIONDISPLAYVALUE);
        rec.OPERATINGUNIT = p[0] || ''; rec.SUB1_OPERATINGUNIT = p[1] || ''; rec.SUB2_OPERATINGUNIT = p[2] || '';
      }
      if (rec.OPERATINGUNIT || rec.SUB2_OPERATINGUNIT) {
        deriveChain_(ctx, rec, UNIT_CHAIN);
        copyUnits_(rec, a);
        ctx.touch('assets', a);
      }
    }
    rec.DEFAULTDIMENSIONDISPLAYVALUE = dimJoin_([rec.OPERATINGUNIT, rec.SUB1_OPERATINGUNIT, rec.SUB2_OPERATINGUNIT]);
    if (rec.ACCUMULATEDDEPRECIATION !== '' && rec.ACCUMULATEDDEPRECIATION < 0) rec.ACCUMULATEDDEPRECIATION = -rec.ACCUMULATEDDEPRECIATION; // D365 يصدّره بالسالب
    if (num_(rec.ACCUMULATEDDEPRECIATION) > num_(rec.ACQUISITIONPRICE)) throw new Error('مجمع الإهلاك أكبر من سعر الاقتناء');
    rec.NETBOOKVALUE = nbv_(rec);
  },

  _lk: function (rec, ex, ctx, entity) {
    const sch = SCHEMAS[entity];
    if (!sch.parentType || !rec.PARENT) return;
    const p = ctx.find(sch.parentType, rec.PARENT);
    if (!p) throw new Error('الكود الأب غير موجود في ' + LKMAP[sch.parentType].ar + ': ' + rec.PARENT);
    rec.PARENT = p.CODE;
  },

  groups: function (rec, ex, ctx) {
    rec.ASSETACCOUNT = postingAccount_(ctx, rec.ASSETACCOUNT, 'حساب الأصل');
    rec.ACCUMDEPACCOUNT = postingAccount_(ctx, rec.ACCUMDEPACCOUNT, 'حساب مجمع الإهلاك');
    rec.DEPEXPENSEACCOUNT = postingAccount_(ctx, rec.DEPEXPENSEACCOUNT, 'حساب مصروف الإهلاك');
    rec.DISPOSALACCOUNT = postingAccount_(ctx, rec.DISPOSALACCOUNT, 'حساب الاستبعاد');
  },

  accounts: function (rec, ex, ctx) {
    const id = low_(rec.MAINACCOUNTID);
    if (rec.ACCOUNTLEVEL === 'Header') {
      rec.PARENTACCOUNT = '';
      if (ex && ex.ACCOUNTLEVEL !== 'Header' && accountUsed_(ctx, id)) throw new Error('لا يمكن تحويل الحساب إلى رئيسي لأنه مستخدم في المجموعات أو الموردين أو اليوميات');
    } else {
      if (ex && ex.ACCOUNTLEVEL === 'Header' && ctx.table('accounts').rows.some(function (r) { return low_(r.PARENTACCOUNT) === id; })) {
        throw new Error('لا يمكن تحويل الحساب إلى فرعي لأن له حسابات فرعية');
      }
      if (rec.PARENTACCOUNT) {
        if (low_(rec.PARENTACCOUNT) === id) throw new Error('الحساب لا يمكن أن يكون أباً لنفسه');
        const p = ctx.find('accounts', rec.PARENTACCOUNT);
        if (!p) throw new Error('الحساب الأب غير موجود: ' + rec.PARENTACCOUNT);
        if (p.ACCOUNTLEVEL !== 'Header') throw new Error('الحساب الأب يجب أن يكون حساباً رئيسياً: ' + rec.PARENTACCOUNT);
        rec.PARENTACCOUNT = p.MAINACCOUNTID;
      }
    }
  },

  vendors: function (rec, ex, ctx) { partyHook_('Vendor')(rec, ex, ctx); },

  journal: function (rec, ex, ctx) {
    if (ex && ex.STATUS && ex.STATUS !== 'Draft') throw new Error('لا يمكن تعديل سطر معتمد أو مُرحّل');
    if (ex) rec.JOURNALBATCHNUMBER = ex.JOURNALBATCHNUMBER;
    if (!rec.JOURNALBATCHNUMBER) rec.JOURNALBATCHNUMBER = ctx.batch();
    if (!ex) {
      const st = ctx.batchStatus(rec.JOURNALBATCHNUMBER);
      if (st && st !== 'Draft') throw new Error('لا يمكن إضافة أسطر إلى يومية معتمدة أو مُرحّلة: ' + rec.JOURNALBATCHNUMBER);
    }
    checkLk_(ctx, rec, 'JOURNALNAME', 'JournalName');
    checkOpenPeriod_(ctx, rec.TRANSDATE, 'سطر يومية الأصول');
    const a = ctx.asset(rec.FIXEDASSETNUMBER);
    if (!a) throw new Error('رقم الأصل غير موجود: ' + rec.FIXEDASSETNUMBER);
    rec.FIXEDASSETNUMBER = a.FIXEDASSETNUMBER;
    rec.ASSETNAME = a.NAME;
    if (!rec.BOOKID) rec.BOOKID = 'General';
    const g = ctx.find('groups', a.FIXEDASSETGROUPID);
    rec.LEDGERACCOUNT = ledgerAccountFor_(rec.TRANSACTIONTYPE, g);
    const d = num_(rec.DEBITAMOUNT), c = num_(rec.CREDITAMOUNT);
    if (d < 0 || c < 0) throw new Error('المبالغ لا يمكن أن تكون سالبة');
    if (d > 0 && c > 0) throw new Error('أدخل مبلغاً في المدين أو الدائن فقط، وليس كليهما');
    if (!d && !c) throw new Error('أدخل مبلغ المدين أو الدائن');
    rec.DEBITAMOUNT = d; rec.CREDITAMOUNT = c;
    if (!rec.OFFSETACCOUNT && rec.OFFSETACCOUNTDISPLAYVALUE) {
      const p = dimSplit_(rec.OFFSETACCOUNTDISPLAYVALUE);
      rec.OFFSETACCOUNT = p[0] || '';
      if (!rec.OFFSET_OPERATINGUNIT && !rec.OFFSET_SUB2_OPERATINGUNIT) {
        rec.OFFSET_OPERATINGUNIT = p[1] || ''; rec.OFFSET_SUB1_OPERATINGUNIT = p[2] || ''; rec.OFFSET_SUB2_OPERATINGUNIT = p[3] || '';
      }
    }
    if (rec.VENDORACCOUNT) {
      const v = ctx.find('vendors', rec.VENDORACCOUNT);
      if (!v) throw new Error('المورد غير موجود: ' + rec.VENDORACCOUNT);
      rec.VENDORACCOUNT = v.VENDORACCOUNTNUMBER;
      if (!rec.OFFSETACCOUNT) rec.OFFSETACCOUNT = v.VENDORLEDGERACCOUNT;
      if (!rec.OFFSETACCOUNT) throw new Error('المورد ' + v.VENDORACCOUNTNUMBER + ' ليس له حساب في دليل الحسابات — حدده من صفحة الموردين');
    }
    if (!rec.OFFSETACCOUNT) rec.OFFSETACCOUNT = defaultOffsetFor_(rec.TRANSACTIONTYPE, g);
    if (!rec.OFFSETACCOUNT) throw new Error('حدد الحساب المقابل من دليل الحسابات');
    rec.OFFSETACCOUNT = postingAccount_(ctx, rec.OFFSETACCOUNT, 'الحساب المقابل');
    if (!rec.OFFSET_OPERATINGUNIT && !rec.OFFSET_SUB1_OPERATINGUNIT && !rec.OFFSET_SUB2_OPERATINGUNIT) copyUnits_(a, rec, 'OFFSET_');
    deriveChain_(ctx, rec, OFFSET_CHAIN);
    rec.OFFSETACCOUNTDISPLAYVALUE = dimJoin_([rec.OFFSETACCOUNT, rec.OFFSET_OPERATINGUNIT, rec.OFFSET_SUB1_OPERATINGUNIT, rec.OFFSET_SUB2_OPERATINGUNIT]);
    if (rec.TRANSACTIONTYPE === 'Depreciation') {
      if (rec.DEPRECIATIONMONTHS === '') rec.DEPRECIATIONMONTHS = 1;
      checkDepDuplicate_(ctx, rec);
    }
    if (!rec.STATUS) rec.STATUS = 'Draft';
  }
};

HOOKS.glheaders = function (rec, ex, ctx) {
  if (ex && ex.STATUS && ex.STATUS !== 'Draft') throw new Error('لا يمكن تعديل قيد معتمد أو مُرحّل');
  if (!rec.VOUCHER) rec.VOUCHER = '@pending'; // الرقم يُعطى عند الكتابة فقط حتى لا تضيع أرقام عند فشل الحفظ
  checkLk_(ctx, rec, 'JOURNALNAME', 'JournalName');
  checkOpenPeriod_(ctx, rec.TRANSDATE, 'القيد');
  ctx.voucherDate = rec.TRANSDATE;
  if (!rec.STATUS) rec.STATUS = 'Draft';
};

HOOKS.customers = partyHook_('Customer');
HOOKS.banks = partyHook_('Bank');

HOOKS.budgets = function (rec, ex, ctx) {
  if (!/^\d{4}$/.test(rec.BUDGETYEAR)) throw new Error('سنة الموازنة يجب أن تكون 4 أرقام (مثل 2026)');
  rec.MAINACCOUNTID = postingAccount_(ctx, rec.MAINACCOUNTID, 'رقم الحساب');
  rec.ACCOUNTNAME = (ctx.find('accounts', rec.MAINACCOUNTID) || {}).NAME || '';
  deriveChain_(ctx, rec, UNIT_CHAIN);
  if (num_(rec.AMOUNT) < 0) throw new Error('المبلغ التقديري لا يمكن أن يكون سالباً');
  const acc = ctx.find('accounts', rec.MAINACCOUNTID);
  const isRev = acc && acc.MAINACCOUNTTYPE === 'Revenue';
  if (rec.BUDGETTYPE === 'Revenue' && !isRev) throw new Error('سطر ' + (rec.LINENUM || '') + ': الحساب ' + rec.MAINACCOUNTID + ' ليس حساب إيرادات (موازنة إيرادات)');
  if (rec.BUDGETTYPE === 'Expense' && isRev) throw new Error('سطر ' + (rec.LINENUM || '') + ': الحساب ' + rec.MAINACCOUNTID + ' حساب إيرادات — استخدم موازنة إيرادات');
  const dup = ctx.table('budgets').rows.filter(function (b) {
    return b.LINEID !== rec.LINEID && low_(b.BUDGETID) !== low_(rec.BUDGETID) && String(b.BUDGETYEAR) === String(rec.BUDGETYEAR) && low_(b.MAINACCOUNTID) === low_(rec.MAINACCOUNTID) &&
      low_(b.OPERATINGUNIT) === low_(rec.OPERATINGUNIT) && low_(b.SUB1_OPERATINGUNIT) === low_(rec.SUB1_OPERATINGUNIT) && low_(b.SUB2_OPERATINGUNIT) === low_(rec.SUB2_OPERATINGUNIT);
  })[0];
  if (dup) throw new Error('سطر ' + (rec.LINENUM || '') + ': الحساب ' + rec.MAINACCOUNTID + ' بنفس الوحدة موجود في الموازنة ' + dup.BUDGETID + ' لنفس السنة');
};

HOOKS.itemgroups = function (rec, ex, ctx) {
  rec.SALESACCOUNT = postingAccount_(ctx, rec.SALESACCOUNT, 'حساب المبيعات');
  rec.PURCHASEACCOUNT = postingAccount_(ctx, rec.PURCHASEACCOUNT, 'حساب المشتريات');
  rec.SALESRETURNACCOUNT = postingAccount_(ctx, rec.SALESRETURNACCOUNT, 'حساب مردودات المبيعات');
  rec.PURCHRETURNACCOUNT = postingAccount_(ctx, rec.PURCHRETURNACCOUNT, 'حساب مردودات المشتريات');
  if (!rec.SALESACCOUNT && !rec.PURCHASEACCOUNT) throw new Error('حدد حساب المبيعات أو حساب المشتريات للمجموعة');
};

function itemHook_(kind) {
  return function (rec, ex, ctx) {
    const g = ctx.find('itemgroups', rec.ITEMGROUP);
    if (!g) throw new Error('مجموعة الصنف غير موجودة "' + (rec.ITEMGROUP || '') + '" — أضفها من صفحة مجموعات الأصناف');
    rec.ITEMGROUP = g.ITEMGROUPID;
    if (kind === 'sales' && !g.SALESACCOUNT) throw new Error('المجموعة ' + g.ITEMGROUPID + ' بدون حساب مبيعات');
    if (kind === 'purchase' && !g.PURCHASEACCOUNT) throw new Error('المجموعة ' + g.ITEMGROUPID + ' بدون حساب مشتريات');
    checkLk_(ctx, rec, 'UNIT', 'ItemUnit');
    if (num_(rec.PRICE) < 0) throw new Error('السعر لا يمكن أن يكون سالباً');
  };
}
HOOKS.salesitems = itemHook_('sales');
HOOKS.purchitems = itemHook_('purchase');

HOOKS.invoices = function (rec, ex, ctx) {
  if (ex && ex.STATUS && ex.STATUS !== 'Draft') throw new Error('لا يمكن تعديل فاتورة مُرحّلة');
  const P = invParty_(rec.INVOICETYPE);
  const p = ctx.find(P.entity, rec.PARTYID);
  if (!p) throw new Error(P.ar + ' غير موجود "' + (rec.PARTYID || '') + '"');
  rec.PARTYID = p[P.id];
  rec.PARTYNAME = p[P.name] || '';
  const local = String(getSettings_().CURRENCY || '').toUpperCase();
  rec.CURRENCYCODE = String(p.CURRENCYCODE || local).toUpperCase();
  if (rec.CURRENCYCODE === local) rec.EXCHANGERATE = 1;
  else if (!(num_(rec.EXCHANGERATE) > 0)) {
    rec.EXCHANGERATE = ctx.rate(rec.CURRENCYCODE, rec.INVOICEDATE);
    if (!(rec.EXCHANGERATE > 0)) throw new Error('لا يوجد سعر صرف للعملة ' + rec.CURRENCYCODE + ' بتاريخ ' + rec.INVOICEDATE);
  }
  deriveChain_(ctx, rec, UNIT_CHAIN);
  if (!rec.SUB2_OPERATINGUNIT) throw new Error('حدد الوحدة الفرعية 2 (الوحدة التي تبيع أو تشتري)');
  if (!rec.DUEDATE) rec.DUEDATE = rec.INVOICEDATE;
  if (String(rec.DUEDATE) < String(rec.INVOICEDATE)) throw new Error('تاريخ الاستحقاق قبل تاريخ الفاتورة');
  checkOpenPeriod_(ctx, rec.INVOICEDATE, 'الفاتورة');
};

/** حساب سطر الفاتورة/الإشعار: المرتجع (إشعار دائن للعميل أو مدين للمورد) يذهب لحساب المردودات
 *  من مجموعة الصنف ثم من الإعدادات، وإلا فحساب المبيعات/المشتريات العادي. */
function lineAccount_(ctx, g) {
  const st = getSettings_(), pur = ctx.invType === 'Purchase';
  const normal = (pur ? g.PURCHASEACCOUNT : g.SALESACCOUNT) || (pur ? st.PURCHASEACCOUNT : st.SALESACCOUNT) || '';
  if (!ctx.invCredit) return (pur ? g.PURCHASEACCOUNT : g.SALESACCOUNT) || (g.ITEMGROUPID ? '' : normal);
  return (pur ? g.PURCHRETURNACCOUNT : g.SALESRETURNACCOUNT) || (pur ? st.PURCHRETURNACCOUNT : st.SALESRETURNACCOUNT) ||
    (pur ? g.PURCHASEACCOUNT : g.SALESACCOUNT) || (g.ITEMGROUPID ? '' : normal);
}

HOOKS.invlines = function (rec, ex, ctx) {
  const lbl = 'سطر ' + (rec.LINENUM || '');
  if (rec.ITEMID) {
    const ent = ctx.invType === 'Purchase' ? 'purchitems' : 'salesitems';
    const it = ctx.find(ent, rec.ITEMID);
    if (!it) throw new Error(lbl + ': الصنف غير موجود في ' + (ent === 'purchitems' ? 'أصناف الشراء' : 'أصناف البيع') + ' "' + rec.ITEMID + '"');
    rec.ITEMID = it.ITEMID;
    if (!rec.ITEMDESCRIPTION) rec.ITEMDESCRIPTION = it.NAME;
    if (!rec.UNIT) rec.UNIT = it.UNIT || '';
    if (!rec.MAINACCOUNTID) {
      const g = ctx.find('itemgroups', it.ITEMGROUP) || {};
      rec.MAINACCOUNTID = lineAccount_(ctx, g);
      if (!rec.MAINACCOUNTID) throw new Error(lbl + ': مجموعة الصنف ' + it.ITEMGROUP + ' بدون حساب ' + (ent === 'purchitems' ? 'مشتريات' : 'مبيعات'));
    }
  }
  if (!rec.MAINACCOUNTID) rec.MAINACCOUNTID = lineAccount_(ctx, {});
  rec.MAINACCOUNTID = postingAccount_(ctx, rec.MAINACCOUNTID, lbl + ' — الحساب');
  rec.ACCOUNTNAME = (ctx.find('accounts', rec.MAINACCOUNTID) || {}).NAME || '';
  deriveChain_(ctx, rec, UNIT_CHAIN);
  if (num_(rec.QTY) <= 0) throw new Error(lbl + ': الكمية يجب أن تكون أكبر من صفر');
  if (num_(rec.UNITPRICE) < 0) throw new Error(lbl + ': السعر لا يمكن أن يكون سالباً');
  if (num_(rec.TAXPERCENT) < 0 || num_(rec.TAXPERCENT) > 100) throw new Error(lbl + ': نسبة الضريبة بين 0 و100');
  rec.LINEAMOUNT = round_(num_(rec.QTY) * num_(rec.UNITPRICE), 3);
  rec.TAXAMOUNT = round_(rec.LINEAMOUNT * num_(rec.TAXPERCENT) / 100, 3);
};

HOOKS.budgetheads = function (rec, ex, ctx) {
  if (!/^\d{4}$/.test(rec.BUDGETYEAR)) throw new Error('سنة الموازنة يجب أن تكون 4 أرقام (مثل 2026)');
  if (!rec.ENTRYDATE) rec.ENTRYDATE = today_();
  if (ex && ex.STATUS && ex.STATUS !== 'Draft') throw new Error('لا يمكن تعديل موازنة معتمدة');
};

HOOKS.cashvouchers = function (rec, ex, ctx) {
  const bank = ctx.find('banks', rec.BANKACCOUNTID);
  if (!bank) throw new Error('الصندوق / البنك غير موجود "' + (rec.BANKACCOUNTID || '') + '"');
  rec.BANKACCOUNTID = bank.BANKACCOUNTID;
  const local = String(getSettings_().CURRENCY || '').toUpperCase();
  rec.CURRENCYCODE = String(bank.CURRENCYCODE || local).toUpperCase();
  if (rec.CURRENCYCODE === local) rec.EXCHANGERATE = 1;
  else if (!(num_(rec.EXCHANGERATE) > 0)) {
    rec.EXCHANGERATE = ctx.rate(rec.CURRENCYCODE, rec.TRANSDATE);
    if (!(rec.EXCHANGERATE > 0)) throw new Error('لا يوجد سعر صرف للعملة ' + rec.CURRENCYCODE + ' بتاريخ ' + rec.TRANSDATE);
  }
  if (!(num_(rec.AMOUNT) > 0)) throw new Error('المبلغ يجب أن يكون أكبر من صفر');
  if (rec.PARTYTYPE === 'Ledger') {
    rec.PARTYID = postingAccount_(ctx, rec.PARTYID, 'الحساب');
    if (!rec.PARTYID) throw new Error('اختر الحساب');
    rec.PARTYNAME = (ctx.find('accounts', rec.PARTYID) || {}).NAME || '';
  } else {
    const P = PARTY[rec.PARTYTYPE];
    if (!P) throw new Error('نوع الطرف غير معروف');
    const p = ctx.find(P.entity, rec.PARTYID);
    if (!p) throw new Error(P.ar + ' غير موجود "' + (rec.PARTYID || '') + '"');
    rec.PARTYID = p[P.id];
    rec.PARTYNAME = p[P.name] || '';
  }
  if (rec.SUB2_OPERATINGUNIT) checkLk_(ctx, rec, 'SUB2_OPERATINGUNIT', 'Sub2_OperatingUnit');
  checkOpenPeriod_(ctx, rec.TRANSDATE, 'السند');
  checkInvoiceLink_(ctx, rec);
};

HOOKS.currencies = function (rec) {
  rec.CURRENCYCODE = String(rec.CURRENCYCODE).toUpperCase();
  if (!/^[A-Z]{3}$/.test(rec.CURRENCYCODE)) throw new Error('رمز العملة يجب أن يكون 3 أحرف إنجليزية (مثل USD)');
};

HOOKS.exrates = function (rec, ex, ctx) {
  const c = ctx.find('currencies', rec.CURRENCYCODE);
  if (!c) throw new Error('العملة غير معرّفة: ' + rec.CURRENCYCODE);
  rec.CURRENCYCODE = c.CURRENCYCODE;
  if (c.CURRENCYCODE === String(getSettings_().CURRENCY).toUpperCase()) throw new Error('لا يلزم سعر صرف للعملة المحلية');
  if (!(num_(rec.RATE) > 0)) throw new Error('سعر الصرف يجب أن يكون أكبر من صفر');
};

HOOKS.fiscal = function (rec, ex, ctx) {
  if (ex && ex.STATUS === 'Closed') throw new Error('لا يمكن تعديل سنة مقفلة — أعد فتحها أولاً');
  if (/^\d{4}$/.test(rec.YEAR)) {
    if (!rec.STARTDATE) rec.STARTDATE = rec.YEAR + '-01-01';
    if (!rec.ENDDATE) rec.ENDDATE = rec.YEAR + '-12-31';
  }
  if (!rec.STARTDATE || !rec.ENDDATE) throw new Error('حدد تاريخ البداية والنهاية');
  if (rec.STARTDATE >= rec.ENDDATE) throw new Error('تاريخ البداية يجب أن يسبق تاريخ النهاية');
  const clash = ctx.table('fiscal').rows.filter(function (r) {
    return r.YEAR && low_(r.YEAR) !== low_(rec.YEAR) && !(String(rec.ENDDATE) < String(r.STARTDATE) || String(rec.STARTDATE) > String(r.ENDDATE));
  })[0];
  if (clash) throw new Error('الفترة تتداخل مع السنة المالية ' + clash.YEAR);
  if (rec.RETAINEDEARNINGSACCOUNT) rec.RETAINEDEARNINGSACCOUNT = postingAccount_(ctx, rec.RETAINEDEARNINGSACCOUNT, 'حساب الأرباح المحتجزة');
  if (!rec.STATUS) rec.STATUS = 'Open';
};

HOOKS.gllines = function (rec, ex, ctx) {
  if (!rec.ACCOUNTTYPE) rec.ACCOUNTTYPE = 'Ledger';
  if (rec.ACCOUNTTYPE !== 'Ledger') {
    const P = PARTY[rec.ACCOUNTTYPE];
    const p = ctx.find(P.entity, rec.PARTYID);
    if (!p) throw new Error('سطر ' + rec.LINENUM + ': ' + P.ar + ' غير موجود "' + (rec.PARTYID || '') + '"');
    rec.PARTYID = p[P.id];
    rec.MAINACCOUNTID = p[P.acc];
    if (!rec.MAINACCOUNTID) throw new Error('سطر ' + rec.LINENUM + ': ' + P.ar + ' ' + rec.PARTYID + ' ليس له حساب في دليل الحسابات');
    if (!rec.CURRENCYCODE) rec.CURRENCYCODE = p.CURRENCYCODE || '';
  } else {
    rec.PARTYID = '';
  }
  rec.MAINACCOUNTID = postingAccount_(ctx, rec.MAINACCOUNTID, 'سطر ' + (rec.LINENUM || '') + ' — رقم الحساب');
  rec.ACCOUNTNAME = (ctx.find('accounts', rec.MAINACCOUNTID) || {}).NAME || '';
  deriveChain_(ctx, rec, UNIT_CHAIN);
  const local = String(getSettings_().CURRENCY || '').toUpperCase();
  rec.CURRENCYCODE = String(rec.CURRENCYCODE || local).toUpperCase();
  // سطور قديمة/تلقائية بدون مبالغ بالعملة: المبلغ المحاسبي هو المبلغ بالعملة المحلية
  if (rec.AMOUNTCURDEBIT === '' && rec.AMOUNTCURCREDIT === '') { rec.AMOUNTCURDEBIT = rec.DEBITAMOUNT; rec.AMOUNTCURCREDIT = rec.CREDITAMOUNT; }
  const d = num_(rec.AMOUNTCURDEBIT), c = num_(rec.AMOUNTCURCREDIT);
  if (d < 0 || c < 0) throw new Error('سطر ' + rec.LINENUM + ': المبالغ لا يمكن أن تكون سالبة');
  if (d > 0 && c > 0) throw new Error('سطر ' + rec.LINENUM + ': أدخل مديناً أو دائناً فقط');
  if (!d && !c) throw new Error('سطر ' + rec.LINENUM + ': أدخل مبلغاً');
  let rate = 1;
  if (rec.CURRENCYCODE !== local) {
    if (!ctx.find('currencies', rec.CURRENCYCODE)) throw new Error('سطر ' + rec.LINENUM + ': العملة غير معرّفة ' + rec.CURRENCYCODE);
    rate = num_(rec.EXCHANGERATE) || ctx.rate(rec.CURRENCYCODE, ctx.voucherDate);
    if (!(rate > 0)) throw new Error('سطر ' + rec.LINENUM + ': لا يوجد سعر صرف للعملة ' + rec.CURRENCYCODE + ' بتاريخ ' + (ctx.voucherDate || '') + ' — أضفه من صفحة أسعار الصرف');
  }
  rec.EXCHANGERATE = rate;
  rec.AMOUNTCURDEBIT = d; rec.AMOUNTCURCREDIT = c;
  rec.DEBITAMOUNT = round_(d * rate, 3); rec.CREDITAMOUNT = round_(c * rate, 3);
  rec.LEDGERDIMENSIONDISPLAYVALUE = dimJoin_([rec.MAINACCOUNTID, rec.OPERATINGUNIT, rec.SUB1_OPERATINGUNIT, rec.SUB2_OPERATINGUNIT]);
};

HOOKS.transfers = function (rec, ex, ctx) {
  if (ex && ex.STATUS && ex.STATUS !== 'Draft') throw new Error('لا يمكن تعديل سطر نقل معتمد أو مُرحّل');
  if (ex) rec.TRANSFERNUMBER = ex.TRANSFERNUMBER;
  if (!rec.TRANSFERNUMBER) rec.TRANSFERNUMBER = ctx.transferNo();
  if (!ex) {
    const st = ctx.docStatus('transfers', 'TRANSFERNUMBER', rec.TRANSFERNUMBER);
    if (st && st !== 'Draft') throw new Error('لا يمكن إضافة أسطر إلى مستند نقل معتمد أو مُرحّل: ' + rec.TRANSFERNUMBER);
  }
  checkOpenPeriod_(ctx, rec.TRANSFERDATE, 'النقل');
  const a = ctx.asset(rec.FIXEDASSETNUMBER);
  if (!a) throw new Error('رقم الأصل غير موجود: ' + rec.FIXEDASSETNUMBER);
  rec.FIXEDASSETNUMBER = a.FIXEDASSETNUMBER;
  rec.ASSETNAME = a.NAME;
  const open = ctx.table('transfers').rows.filter(function (r) {
    return low_(r.FIXEDASSETNUMBER) === low_(a.FIXEDASSETNUMBER) && r.STATUS !== 'Posted' && r.LINEID !== rec.LINEID;
  })[0];
  if (open) throw new Error('الأصل ' + a.FIXEDASSETNUMBER + ' له نقل غير مُرحّل في المستند ' + open.TRANSFERNUMBER);
  TRF_FIELDS.forEach(function (k) { rec['FROM_' + k] = a[k] || ''; });
  deriveChain_(ctx, rec, TO_UNIT_CHAIN);
  deriveChain_(ctx, rec, TO_LOCATION_CHAIN);
  checkLk_(ctx, rec, 'TO_DEPARTMENT', 'Department');
  checkLk_(ctx, rec, 'TO_RESPONSIBLE', 'Employee');
  if (TRF_FIELDS.every(function (k) { return low_(rec['TO_' + k]) === low_(rec['FROM_' + k]); })) {
    throw new Error('لا يوجد تغيير: الوجهة مطابقة للوضع الحالي للأصل ' + a.FIXEDASSETNUMBER);
  }
  if (!rec.STATUS) rec.STATUS = 'Draft';
};

/** منع تكرار الإهلاك: لا يُقبل إهلاك لفترة سبق إهلاكها، ولا سطر إهلاك ثانٍ غير مُرحّل لنفس الأصل والدفتر. */
function checkDepDuplicate_(ctx, rec) {
  const b = ctx.table('books').rows.filter(function (x) {
    return low_(x.FIXEDASSETNUMBER) === low_(rec.FIXEDASSETNUMBER) && low_(x.BOOKID) === low_(rec.BOOKID || 'General');
  })[0];
  if (b && b.LASTDEPRECIATIONDATE && ym_(rec.TRANSDATE) <= ym_(b.LASTDEPRECIATIONDATE)) {
    throw new Error('تم احتساب إهلاك الأصل ' + rec.FIXEDASSETNUMBER + ' حتى ' + b.LASTDEPRECIATIONDATE + ' — لا يمكن تكرار الإهلاك لنفس الفترة');
  }
  const other = ctx.table('journal').rows.filter(function (j) {
    return j.TRANSACTIONTYPE === 'Depreciation' && j.STATUS !== 'Posted' && j.LINEID !== rec.LINEID &&
      low_(j.FIXEDASSETNUMBER) === low_(rec.FIXEDASSETNUMBER) && low_(j.BOOKID || 'General') === low_(rec.BOOKID || 'General');
  })[0];
  if (other) throw new Error('يوجد إهلاك غير مُرحّل للأصل ' + rec.FIXEDASSETNUMBER + ' في اليومية ' + other.JOURNALBATCHNUMBER + ' — رحّله أو احذفه أولاً');
}

/** بعد حفظ الأصل: تحديث أبعاد دفاتره (الأبعاد على الدفتر نسخة من الأصل). */
const AFTER = {
  assets: function (rec, ctx) {
    const bt = ctx.table('books');
    const id = low_(rec.FIXEDASSETNUMBER);
    bt.rows.forEach(function (b, i) {
      if (low_(b.FIXEDASSETNUMBER) !== id) return;
      if (b.OPERATINGUNIT === rec.OPERATINGUNIT && b.SUB1_OPERATINGUNIT === rec.SUB1_OPERATINGUNIT && b.SUB2_OPERATINGUNIT === rec.SUB2_OPERATINGUNIT) return;
      copyUnits_(rec, b);
      b.DEFAULTDIMENSIONDISPLAYVALUE = dimJoin_([b.OPERATINGUNIT, b.SUB1_OPERATINGUNIT, b.SUB2_OPERATINGUNIT]);
      writeRow_(bt, i, b);
    });
  }
};

function accountUsed_(ctx, id) {
  return ctx.table('groups').rows.some(function (g) {
    return [g.ASSETACCOUNT, g.ACCUMDEPACCOUNT, g.DEPEXPENSEACCOUNT, g.DISPOSALACCOUNT].some(function (x) { return low_(x) === id; });
  }) || ctx.table('vendors').rows.some(function (v) { return low_(v.VENDORLEDGERACCOUNT) === id; }) ||
    ctx.table('journal').rows.some(function (j) { return low_(j.OFFSETACCOUNT) === id || low_(j.LEDGERACCOUNT) === id; });
}

/** يجهّز سجلاً للحفظ. mode: add | edit | upsert */
function prepare_(entity, raw, mode, ctx) {
  const sch = SCHEMAS[entity];
  const t = ctx.table(entity);
  const rec = clean_(sch, raw, ctx.dateFmt);
  if (sch.autoId && mode === 'add') rec[sch.autoId] = '';
  const hasKey = sch.key.every(function (k) { return rec[k] !== ''; });
  const found = hasKey ? ctx.index(entity)[keyOf_(sch, rec)] : undefined;
  const idx = found === undefined ? -1 : found;
  if (mode === 'add' && idx >= 0) throw new Error('السجل موجود مسبقاً: ' + sch.key.map(function (k) { return rec[k]; }).join(' / '));
  if (mode === 'edit' && idx < 0) throw new Error('السجل غير موجود');
  const ex = idx >= 0 ? t.rows[idx] : null;
  if (!ex && sch.autoId) rec[sch.autoId] = 'L' + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 6).toUpperCase();
  sch.fields.forEach(function (f) {
    if (f.sys && !f.key) rec[f.k] = ex ? ex[f.k] : '';
    if (!ex && rec[f.k] === '' && f.dflt !== undefined) rec[f.k] = f.dflt === '@today' ? today_() : f.dflt;
  });
  const hook = HOOKS[entity] || (sch.lkType ? HOOKS._lk : null);
  if (hook) hook(rec, ex, ctx, entity);
  sch.fields.forEach(function (f) {
    if (f.req && rec[f.k] === '') throw new Error('الحقل مطلوب: ' + f.ar + ' (' + f.en + ')');
  });
  const now = stamp_();
  rec.CREATEDBY = ex ? ex.CREATEDBY : ctx.s.username;
  rec.CREATEDON = ex ? ex.CREATEDON : now;
  rec.MODIFIEDBY = ex ? ctx.s.username : '';
  rec.MODIFIEDON = ex ? now : '';
  return { idx: idx, rec: rec, isNew: !ex };
}

function commit_(entity, p, ctx) {
  const t = ctx.table(entity);
  if (p.idx >= 0) writeRow_(t, p.idx, p.rec);
  else {
    const i = appendRow_(t, p.rec);
    ctx.index(entity)[keyOf_(SCHEMAS[entity], p.rec)] = i;
  }
  if (AFTER[entity]) AFTER[entity](p.rec, ctx);
}

/* ============================== CRUD API ============================== */

function schema_(entity) {
  const sch = SCHEMAS[entity];
  if (!sch) throw new Error('كيان غير معروف: ' + entity);
  return sch;
}

function listRecords(token, entity) {
  const s = auth_(token);
  const sch = schema_(entity);
  need_(s, sch.perm, 'view');
  return readTable_(sch.sheet).rows.filter(function (r) { return !blank_(r); });
}

function saveRecord(token, entity, rec, isNew) {
  const s = auth_(token);
  const sch = schema_(entity);
  need_(s, sch.perm, isNew ? 'add' : 'edit');
  return withLock_(function () {
    const ctx = ctx_(s);
    const p = prepare_(entity, rec, isNew ? 'add' : 'edit', ctx);
    commit_(entity, p, ctx);
    ctx.flush();
    return p.rec;
  });
}

function deleteRecords(token, entity, keys) {
  const s = auth_(token);
  const sch = schema_(entity);
  need_(s, sch.perm, 'delete');
  return withLock_(function () {
    const ctx = ctx_(s);
    const t = ctx.table(entity);
    const idxs = [];
    (keys || []).forEach(function (k) {
      const i = ctx.index(entity)[keyOf_(sch, k)];
      if (i === undefined) return;
      const r = t.rows[i];
      if ((entity === 'journal' || entity === 'transfers') && r.STATUS !== 'Draft') throw new Error('لا يمكن حذف سطر معتمد أو مُرحّل');
      if (entity === 'assets') {
        const n = low_(r.FIXEDASSETNUMBER);
        if (ctx.table('books').rows.some(function (b) { return low_(b.FIXEDASSETNUMBER) === n; }) ||
            ctx.table('journal').rows.some(function (j) { return low_(j.FIXEDASSETNUMBER) === n; }) ||
            ctx.table('transfers').rows.some(function (j) { return low_(j.FIXEDASSETNUMBER) === n; })) {
          throw new Error('لا يمكن حذف الأصل ' + r.FIXEDASSETNUMBER + ' لوجود دفاتر أو معاملات مرتبطة به');
        }
      }
      if (entity === 'books') {
        const n = low_(r.FIXEDASSETNUMBER), b = low_(r.BOOKID);
        if (ctx.table('journal').rows.some(function (j) { return low_(j.FIXEDASSETNUMBER) === n && low_(j.BOOKID) === b && j.STATUS === 'Posted'; })) {
          throw new Error('لا يمكن حذف دفتر له معاملات مُرحّلة');
        }
      }
      if (entity === 'vendors') {
        const v = low_(r.VENDORACCOUNTNUMBER);
        if (ctx.table('journal').rows.some(function (j) { return low_(j.VENDORACCOUNT) === v; })) {
          throw new Error('لا يمكن حذف المورد ' + r.VENDORACCOUNTNUMBER + ' لوجود معاملات مرتبطة به');
        }
      }
      if (sch.lkType) {
        const code = low_(r.CODE);
        const child = LK.filter(function (t) { return t.parent === sch.lkType; })[0];
        if (child && ctx.table(child.v).rows.some(function (x) { return low_(x.PARENT) === code; })) {
          throw new Error('لا يمكن حذف ' + r.CODE + ' لأن له عناصر تابعة في ' + child.ar);
        }
      }
      if (entity === 'customers' || entity === 'banks' || entity === 'vendors') {
        const T = entity === 'customers' ? 'Customer' : entity === 'banks' ? 'Bank' : 'Vendor';
        const id = r[PARTY[T].id];
        if (partyUsed_(ctx, T, id)) throw new Error('لا يمكن حذف ' + PARTY[T].ar + ' ' + id + ' لوجود قيود مرتبطة به');
        if (T !== 'Bank' && ctx.table('invoices').rows.some(function (v) { return invParty_(v.INVOICETYPE) === PARTY[T] && low_(v.PARTYID) === low_(id); })) {
          throw new Error('لا يمكن حذف ' + PARTY[T].ar + ' ' + id + ' لوجود فواتير مرتبطة به');
        }
        if (ctx.table('cashvouchers').rows.some(function (c) { return low_(T === 'Bank' ? c.BANKACCOUNTID : c.PARTYTYPE === T ? c.PARTYID : '') === low_(id); })) {
          throw new Error('لا يمكن حذف ' + PARTY[T].ar + ' ' + id + ' لوجود سندات قبض/صرف مرتبطة به');
        }
      }
      if (entity === 'itemgroups') {
        const ig = low_(r.ITEMGROUPID);
        if (['salesitems', 'purchitems'].some(function (e2) { return ctx.table(e2).rows.some(function (x) { return low_(x.ITEMGROUP) === ig; }); })) {
          throw new Error('لا يمكن حذف المجموعة ' + r.ITEMGROUPID + ' لوجود أصناف مرتبطة بها');
        }
      }
      if (entity === 'salesitems' || entity === 'purchitems') {
        const iid = low_(r.ITEMID), it = entity === 'salesitems' ? 'Sales' : 'Purchase';
        const invT = {};
        ctx.table('invoices').rows.forEach(function (h) { invT[low_(h.INVOICEID)] = h.INVOICETYPE; });
        if (ctx.table('invlines').rows.some(function (l) { return low_(l.ITEMID) === iid && invT[low_(l.INVOICEID)] === it; })) {
          throw new Error('لا يمكن حذف الصنف ' + r.ITEMID + ' لوجود فواتير مرتبطة به');
        }
      }
      if (entity === 'groups') {
        const gid = low_(r.FIXEDASSETGROUPID);
        if (ctx.table('assets').rows.some(function (a) { return low_(a.FIXEDASSETGROUPID) === gid; })) {
          throw new Error('لا يمكن حذف المجموعة ' + r.FIXEDASSETGROUPID + ' لوجود أصول مرتبطة بها');
        }
      }
      if (entity === 'accounts') {
        const id = low_(r.MAINACCOUNTID);
        if (t.rows.some(function (x) { return low_(x.PARENTACCOUNT) === id; })) throw new Error('لا يمكن حذف الحساب ' + r.MAINACCOUNTID + ' لأن له حسابات فرعية');
        if (accountUsed_(ctx, id)) throw new Error('لا يمكن حذف الحساب ' + r.MAINACCOUNTID + ' لأنه مستخدم في المجموعات أو الموردين أو اليوميات');
      }
      idxs.push(i);
    });
    deleteRows_(t, idxs);
    return idxs.length;
  });
}

/** استيراد/تحديث مجموعة سجلات (من ملفات Excel/CSV المصدّرة من Dynamics 365). */
function importRecords(token, entity, rows, opts) {
  const s = auth_(token);
  const sch = schema_(entity);
  need_(s, sch.perm, 'view');
  opts = opts || {};
  if (!rows || !rows.length) return { added: 0, updated: 0, errors: [] };
  if (rows.length > 2000) throw new Error('الحد الأقصى 2000 سطر في الدفعة الواحدة');
  return withLock_(function () {
    const ctx = ctx_(s);
    if (entity === 'journal' && opts.batch) ctx.importBatch = opts.batch;
    const t = ctx.table(entity);
    const res = { added: 0, updated: 0, errors: [], batch: '' };
    const canAdd = can_(s, sch.perm, 'add'), canEdit = can_(s, sch.perm, 'edit');
    const startLen = t.rows.length;
    const changed = {};
    const done = [];
    let order = rows.map(function (r, n) { return { raw: r, n: n }; });
    if (entity === 'accounts') {
      // الحسابات الرئيسية أولاً حتى تجد الحسابات الفرعية آباءها
      order = order.filter(function (x) { return norm_(x.raw.ACCOUNTLEVEL) === 'header' || norm_(x.raw.ACCOUNTLEVEL) === 'رئيسي'; })
        .concat(order.filter(function (x) { return !(norm_(x.raw.ACCOUNTLEVEL) === 'header' || norm_(x.raw.ACCOUNTLEVEL) === 'رئيسي'); }));
    }
    order.forEach(function (o) {
      const raw = o.raw, n = o.n;
      try {
        const p = prepare_(entity, raw, 'upsert', ctx);
        if (p.isNew && !canAdd) throw new Error('ليس لديك صلاحية الإضافة');
        if (!p.isNew && !canEdit) throw new Error('ليس لديك صلاحية التعديل');
        if (p.isNew) {
          t.rows.push(p.rec);
          ctx.index(entity)[keyOf_(sch, p.rec)] = t.rows.length - 1;
          res.added++;
        } else {
          t.rows[p.idx] = p.rec;
          changed[p.idx] = 1;
          res.updated++;
        }
        done.push(p.rec);
      } catch (e) {
        res.errors.push({ row: (opts.offset || 0) + n + 1, msg: e.message });
      }
    });
    Object.keys(changed).forEach(function (i) {
      t.sh.getRange(Number(i) + 2, 1, 1, t.head.length).setValues([toRow_(t.head, t.rows[i])]);
    });
    const added = t.rows.slice(startLen);
    if (added.length) {
      const need = startLen + 1 + added.length;
      if (need > t.sh.getMaxRows()) t.sh.insertRowsAfter(t.sh.getMaxRows(), need - t.sh.getMaxRows() + 100);
      t.sh.getRange(startLen + 2, 1, added.length, t.head.length).setValues(added.map(function (r) { return toRow_(t.head, r); }));
    }
    if (AFTER[entity]) done.forEach(function (r) { AFTER[entity](r, ctx); });
    ctx.flush();
    res.errors.sort(function (a, b) { return a.row - b.row; });
    res.batch = ctx.importBatch;
    return res;
  });
}

/* ============================== Journal workflow ============================== */

function newJournalBatch(token) {
  const s = auth_(token);
  need_(s, 'journal', 'add');
  return withLock_(nextBatch_);
}

/** action: approve | unapprove | post */
function journalAction(token, batch, action) {
  const s = auth_(token);
  need_(s, 'journal', action === 'post' ? 'post' : 'approve');
  return withLock_(function () {
    const ctx = ctx_(s);
    const t = ctx.table('journal');
    const idxs = [];
    t.rows.forEach(function (r, i) { if (low_(r.JOURNALBATCHNUMBER) === low_(batch)) idxs.push(i); });
    if (!idxs.length) throw new Error('اليومية غير موجودة أو لا تحتوي على أسطر');
    const now = stamp_();
    const lines = idxs.map(function (i) { return t.rows[i]; });

    if (action === 'approve') {
      if (lines.some(function (l) { return l.STATUS !== 'Draft'; })) throw new Error('اليومية ليست في حالة مسودة');
      lines.forEach(function (l) { l.STATUS = 'Approved'; l.APPROVEDBY = s.username; l.APPROVEDON = now; });
    } else if (action === 'unapprove') {
      if (lines.some(function (l) { return l.STATUS !== 'Approved'; })) throw new Error('يمكن إلغاء الاعتماد لليوميات المعتمدة غير المرحّلة فقط');
      lines.forEach(function (l) { l.STATUS = 'Draft'; l.APPROVEDBY = ''; l.APPROVEDON = ''; });
    } else if (action === 'post') {
      if (lines.some(function (l) { return l.STATUS === 'Posted'; })) throw new Error('اليومية مُرحّلة مسبقاً');
      if (lines.some(function (l) { return l.STATUS !== 'Approved'; })) throw new Error('يجب اعتماد اليومية قبل الترحيل');
      lines.forEach(function (l) { checkOpenPeriod_(ctx, l.TRANSDATE, 'سطر يومية الأصول'); });
      const legs = [];
      lines.forEach(function (l) { faLegs_(ctx, l).forEach(function (x) { legs.push(x); }); });
      const v = autoVoucher_(ctx, { TRANSDATE: lines[0].TRANSDATE, JOURNALNAME: lines[0].JOURNALNAME, DESCRIPTION: 'قيد يومية الأصول ' + batch,
        REFERENCE: batch, SOURCE: 'FAJournal', SOURCEREF: batch }, legs);
      postToBooks_(lines, ctx);
      writeVoucher_(ctx, v);
      lines.forEach(function (l) { l.STATUS = 'Posted'; l.POSTEDBY = s.username; l.POSTEDON = now; l.GLVOUCHER = v.head.rec.VOUCHER; });
    } else {
      throw new Error('إجراء غير معروف');
    }
    idxs.forEach(function (i) { writeRow_(t, i, t.rows[i]); });
    return lines.length;
  });
}

function postToBooks_(lines, ctx) {
  const bt = ctx.table('books');
  const bi = ctx.index('books');
  const dirty = {};
  lines.slice().sort(function (a, b) { return String(a.TRANSDATE).localeCompare(String(b.TRANSDATE)); }).forEach(function (l) {
    const i = bi[low_(l.FIXEDASSETNUMBER) + '|' + low_(l.BOOKID || 'General')];
    if (i === undefined) throw new Error('لا يوجد دفتر "' + (l.BOOKID || 'General') + '" للأصل ' + l.FIXEDASSETNUMBER + ' — أنشئ الدفتر أولاً');
    const b = bt.rows[i];
    const amt = num_(l.DEBITAMOUNT) - num_(l.CREDITAMOUNT);
    switch (l.TRANSACTIONTYPE) {
      case 'Acquisition':
      case 'AcquisitionAdj':
        b.ACQUISITIONPRICE = round_(num_(b.ACQUISITIONPRICE) + amt, 3);
        if (l.TRANSACTIONTYPE === 'Acquisition' && !b.ACQUISITIONDATE) b.ACQUISITIONDATE = l.TRANSDATE;
        if (!b.DEPRECIATIONSTARTDATE && l.TRANSACTIONTYPE === 'Acquisition') b.DEPRECIATIONSTARTDATE = l.TRANSDATE;
        if (!b.STATUS || b.STATUS === 'NoAcquisition') b.STATUS = 'Open';
        break;
      case 'Depreciation':
        if (b.LASTDEPRECIATIONDATE && ym_(l.TRANSDATE) <= ym_(b.LASTDEPRECIATIONDATE)) {
          throw new Error('تكرار إهلاك: الأصل ' + l.FIXEDASSETNUMBER + ' تم إهلاكه حتى ' + b.LASTDEPRECIATIONDATE + ' ولا يمكن ترحيل إهلاك بتاريخ ' + l.TRANSDATE);
        }
        b.LASTDEPRECIATIONDATE = l.TRANSDATE;
        if (b.DEPRECIATIONPERIODSREMAINING !== '') {
          b.DEPRECIATIONPERIODSREMAINING = Math.max(0, num_(b.DEPRECIATIONPERIODSREMAINING) - (num_(l.DEPRECIATIONMONTHS) || 1));
        }
        b.ACCUMULATEDDEPRECIATION = round_(num_(b.ACCUMULATEDDEPRECIATION) - amt, 3);
        break;
      case 'DepreciationAdj':
      case 'ExtraordinaryDepreciation':
        b.ACCUMULATEDDEPRECIATION = round_(num_(b.ACCUMULATEDDEPRECIATION) - amt, 3);
        break;
      case 'DisposalSale': b.STATUS = 'Sold'; break;
      case 'DisposalScrap': b.STATUS = 'Scrapped'; break;
      default: break;
    }
    b.NETBOOKVALUE = nbv_(b);
    b.MODIFIEDBY = ctx.s.username; b.MODIFIEDON = stamp_();
    dirty[i] = 1;
  });
  Object.keys(dirty).forEach(function (i) { writeRow_(bt, Number(i), bt.rows[i]); });
}

/* ============================== Helpers: books & depreciation ============================== */

function nbv_(b) { return round_(num_(b.ACQUISITIONPRICE) - num_(b.ACCUMULATEDDEPRECIATION), 3); }
function ym_(d) { const m = String(d || '').match(/^(\d{4})-(\d{2})/); return m ? Number(m[1]) * 12 + Number(m[2]) - 1 : null; }
function ymStr_(n) { return Math.floor(n / 12) + '-' + ('0' + (n % 12 + 1)).slice(-2); }
function monthEnd_(d) {
  const n = ym_(d);
  const y = Math.floor(n / 12), m = n % 12 + 1;
  const last = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return y + '-' + ('0' + m).slice(-2) + '-' + ('0' + last).slice(-2);
}

/** الدفتر الرئيسي لكل أصل (General أولاً) — لتجنّب تكرار القيمة عند وجود أكثر من دفتر. */
function primaryBooks_(books) {
  const m = {};
  books.forEach(function (b) {
    if (!b.FIXEDASSETNUMBER) return;
    const k = low_(b.FIXEDASSETNUMBER);
    if (!m[k] || (low_(b.BOOKID) === 'general' && low_(m[k].BOOKID) !== 'general')) m[k] = b;
  });
  return m;
}

/* ============================== Asset transfers ============================== */

function newTransferNumber(token) {
  const s = auth_(token);
  need_(s, 'transfers', 'add');
  return withLock_(function () { return nextSeq_('TRANSFERSEQ', 'TRANSFERPREFIX', 'TRF-'); });
}

/** نقل جماعي: ينشئ مستند نقل (مسودة) لعدة أصول إلى نفس الوجهة. الحقول الفارغة في الوجهة = بدون تغيير. */
function createTransferDocument(token, p) {
  const s = auth_(token);
  need_(s, 'transfers', 'add');
  p = p || {};
  const ids = (p.assets || []).filter(String);
  if (!ids.length) throw new Error('حدد أصلاً واحداً على الأقل');
  if (ids.length > 500) throw new Error('الحد الأقصى 500 أصل في المستند الواحد');
  const to = p.to || {};
  const unitGiven = TRF_UNIT.some(function (k) { return to['TO_' + k]; });
  const locGiven = TRF_LOC.some(function (k) { return to['TO_' + k]; });
  if (!unitGiven && !locGiven && !to.TO_DEPARTMENT && !to.TO_RESPONSIBLE) throw new Error('حدد الوجهة الجديدة (وحدة أو موقع أو إدارة أو موظف)');
  return withLock_(function () {
    const ctx = ctx_(s);
    ctx.trfNo = p.number || '';
    const res = { number: '', count: 0, skipped: [] };
    ids.forEach(function (id) {
      try {
        const a = ctx.asset(id);
        if (!a) throw new Error('الأصل غير موجود');
        const rec = { TRANSFERNUMBER: ctx.transferNo(), TRANSFERDATE: p.TRANSFERDATE || '', REASON: p.REASON || '', FIXEDASSETNUMBER: a.FIXEDASSETNUMBER };
        TRF_FIELDS.forEach(function (k) { rec['TO_' + k] = a[k] || ''; });
        if (unitGiven) TRF_UNIT.forEach(function (k) { rec['TO_' + k] = to['TO_' + k] || ''; });
        if (locGiven) TRF_LOC.forEach(function (k) { rec['TO_' + k] = to['TO_' + k] || ''; });
        if (to.TO_DEPARTMENT) rec.TO_DEPARTMENT = to.TO_DEPARTMENT;
        if (to.TO_RESPONSIBLE) rec.TO_RESPONSIBLE = to.TO_RESPONSIBLE;
        commit_('transfers', prepare_('transfers', rec, 'add', ctx), ctx);
        res.count++;
      } catch (e) {
        res.skipped.push({ asset: id, msg: e.message });
      }
    });
    if (!res.count) throw new Error(res.skipped.map(function (x) { return x.asset + ': ' + x.msg; }).join(' | '));
    res.number = ctx.trfNo;
    return res;
  });
}

/** action: approve | unapprove | post — الترحيل يطبّق الوجهة الجديدة على الأصل ودفاتره. */
function transferAction(token, number, action) {
  const s = auth_(token);
  need_(s, 'transfers', action === 'post' ? 'post' : 'approve');
  return withLock_(function () {
    const ctx = ctx_(s);
    const t = ctx.table('transfers');
    const idxs = [];
    t.rows.forEach(function (r, i) { if (low_(r.TRANSFERNUMBER) === low_(number)) idxs.push(i); });
    if (!idxs.length) throw new Error('مستند النقل غير موجود أو لا يحتوي على أسطر');
    const now = stamp_();
    const lines = idxs.map(function (i) { return t.rows[i]; });
    if (action === 'approve') {
      if (lines.some(function (l) { return l.STATUS !== 'Draft'; })) throw new Error('المستند ليس في حالة مسودة');
      lines.forEach(function (l) { l.STATUS = 'Approved'; l.APPROVEDBY = s.username; l.APPROVEDON = now; });
    } else if (action === 'unapprove') {
      if (lines.some(function (l) { return l.STATUS !== 'Approved'; })) throw new Error('يمكن إلغاء الاعتماد للمستندات المعتمدة غير المرحّلة فقط');
      lines.forEach(function (l) { l.STATUS = 'Draft'; l.APPROVEDBY = ''; l.APPROVEDON = ''; });
    } else if (action === 'post') {
      if (lines.some(function (l) { return l.STATUS === 'Posted'; })) throw new Error('المستند مُرحّل مسبقاً');
      if (lines.some(function (l) { return l.STATUS !== 'Approved'; })) throw new Error('يجب اعتماد المستند قبل الترحيل');
      lines.forEach(function (l) { checkOpenPeriod_(ctx, l.TRANSFERDATE, 'النقل'); });
      const at = ctx.table('assets');
      lines.forEach(function (l) {
        const a = ctx.asset(l.FIXEDASSETNUMBER);
        if (!a) throw new Error('الأصل غير موجود: ' + l.FIXEDASSETNUMBER);
        TRF_FIELDS.forEach(function (k) {
          if (low_(a[k]) !== low_(l['FROM_' + k])) {
            throw new Error('الأصل ' + a.FIXEDASSETNUMBER + ' تغيّرت بياناته بعد إنشاء النقل — ألغِ الاعتماد واحفظ السطر لتحديث "من"');
          }
        });
      });
      const legs = [];
      lines.forEach(function (l) { transferLegs_(ctx, l).forEach(function (x) { legs.push(x); }); });
      const st = getSettings_();
      const v = legs.length ? autoVoucher_(ctx, { TRANSDATE: lines[0].TRANSFERDATE, JOURNALNAME: st.TRFJOURNALNAME || 'FAJ',
        DESCRIPTION: 'قيد نقل أصول بين الوحدات ' + number, REFERENCE: number, SOURCE: 'Transfer', SOURCEREF: number }, legs) : null;
      if (v) writeVoucher_(ctx, v);
      lines.forEach(function (l) {
        const a = ctx.asset(l.FIXEDASSETNUMBER);
        if (v) l.GLVOUCHER = v.head.rec.VOUCHER;
        TRF_FIELDS.forEach(function (k) { a[k] = l['TO_' + k] || ''; });
        a.MODIFIEDBY = s.username; a.MODIFIEDON = now;
        writeRow_(at, at.rows.indexOf(a), a);
        AFTER.assets(a, ctx);
        l.STATUS = 'Posted'; l.POSTEDBY = s.username; l.POSTEDON = now;
      });
    } else {
      throw new Error('إجراء غير معروف');
    }
    idxs.forEach(function (i) { writeRow_(t, i, t.rows[i]); });
    return lines.length;
  });
}

/* ============================== General ledger vouchers ============================== */

/**
 * أطراف القيد لسطر يومية أصول: حساب الأصل (من المجموعة) والحساب المقابل هو الطرف الثاني.
 * مبلغ مدين في يومية الأصول = مدين حساب الأصل / دائن الحساب المقابل، والعكس.
 */
function faLegs_(ctx, l) {
  const a = ctx.asset(l.FIXEDASSETNUMBER);
  if (!a) throw new Error('الأصل غير موجود: ' + l.FIXEDASSETNUMBER);
  const g = ctx.find('groups', a.FIXEDASSETGROUPID);
  const ledger = ledgerAccountFor_(l.TRANSACTIONTYPE, g);
  if (!ledger) throw new Error('مجموعة الأصل ' + a.FIXEDASSETGROUPID + ' ليس لها ' + (DEP_TYPES[l.TRANSACTIONTYPE] ? 'حساب مجمع إهلاك' : 'حساب أصل') + ' — حدده في صفحة مجموعات الأصول');
  if (!l.OFFSETACCOUNT) throw new Error('سطر الأصل ' + l.FIXEDASSETNUMBER + ' بدون حساب مقابل');
  l.LEDGERACCOUNT = ledger;
  const amt = round_(num_(l.DEBITAMOUNT) - num_(l.CREDITAMOUNT), 3);
  const assetSide = { acc: ledger, ou: a.OPERATINGUNIT, s1: a.SUB1_OPERATINGUNIT, s2: a.SUB2_OPERATINGUNIT };
  const offSide = { acc: l.OFFSETACCOUNT, ou: l.OFFSET_OPERATINGUNIT, s1: l.OFFSET_SUB1_OPERATINGUNIT, s2: l.OFFSET_SUB2_OPERATINGUNIT };
  if (l.VENDORACCOUNT) { const v = ctx.find('vendors', l.VENDORACCOUNT); if (v && low_(v.VENDORLEDGERACCOUNT) === low_(l.OFFSETACCOUNT)) { offSide.pt = 'Vendor'; offSide.pid = v.VENDORACCOUNTNUMBER; } }
  const note = (l.TEXT || '') + (l.TEXT ? ' — ' : '') + l.FIXEDASSETNUMBER;
  const leg = function (side, dr, cr) { return Object.assign({ debit: dr, credit: cr, note: note, ref: l.FIXEDASSETNUMBER }, side); };
  return amt >= 0 ? [leg(assetSide, amt, 0), leg(offSide, 0, amt)] : [leg(offSide, -amt, 0), leg(assetSide, 0, -amt)];
}

/** قيد إعادة التبويب عند نقل الأصل بين الوحدات: التكلفة ومجمع الإهلاك تنتقل من الوحدات القديمة إلى الجديدة. */
function transferLegs_(ctx, l) {
  if (TRF_UNIT.every(function (k) { return low_(l['FROM_' + k]) === low_(l['TO_' + k]); })) return [];
  const a = ctx.asset(l.FIXEDASSETNUMBER);
  const b = primaryBooks_(ctx.table('books').rows)[low_(l.FIXEDASSETNUMBER)];
  if (!b) return [];
  const cost = round_(num_(b.ACQUISITIONPRICE), 3), accum = round_(num_(b.ACCUMULATEDDEPRECIATION), 3);
  if (!cost && !accum) return [];
  const g = ctx.find('groups', a.FIXEDASSETGROUPID);
  if (cost && !(g && g.ASSETACCOUNT)) throw new Error('مجموعة الأصل ' + a.FIXEDASSETGROUPID + ' ليس لها حساب أصل — لا يمكن إنشاء قيد النقل');
  if (accum && !(g && g.ACCUMDEPACCOUNT)) throw new Error('مجموعة الأصل ' + a.FIXEDASSETGROUPID + ' ليس لها حساب مجمع إهلاك — لا يمكن إنشاء قيد النقل');
  const from = { ou: l.FROM_OPERATINGUNIT, s1: l.FROM_SUB1_OPERATINGUNIT, s2: l.FROM_SUB2_OPERATINGUNIT };
  const to = { ou: l.TO_OPERATINGUNIT, s1: l.TO_SUB1_OPERATINGUNIT, s2: l.TO_SUB2_OPERATINGUNIT };
  const note = 'نقل ' + a.FIXEDASSETNUMBER + ' بين الوحدات';
  const legs = [];
  const leg = function (acc, units, dr, cr, what) { legs.push(Object.assign({ acc: acc, debit: dr, credit: cr, note: note + ' (' + what + ')', ref: a.FIXEDASSETNUMBER }, units)); };
  if (cost) { leg(g.ASSETACCOUNT, to, cost, 0, 'التكلفة'); leg(g.ASSETACCOUNT, from, 0, cost, 'التكلفة'); }
  if (accum) { leg(g.ACCUMDEPACCOUNT, from, accum, 0, 'مجمع الإهلاك'); leg(g.ACCUMDEPACCOUNT, to, 0, accum, 'مجمع الإهلاك'); }
  return legs;
}

/** يجهّز قيداً (رأس + سطور) ويتحقق منه دون كتابة. */
function prepareVoucher_(ctx, head, legs, mode) {
  const t = ctx.table('glheaders');
  if (!legs.length) throw new Error('القيد لا يحتوي على سطور');
  const ph = prepare_('glheaders', head, mode, ctx);
  const lines = legs.map(function (x, i) {
    return prepare_('gllines', {
      VOUCHER: ph.rec.VOUCHER, LINENUM: i + 1, ACCOUNTTYPE: x.pt || 'Ledger', PARTYID: x.pid || '', MAINACCOUNTID: x.acc || '', OPERATINGUNIT: x.ou || '', SUB1_OPERATINGUNIT: x.s1 || '',
      SUB2_OPERATINGUNIT: x.s2 || '', CURRENCYCODE: x.cur || '', EXCHANGERATE: x.rate || '',
      AMOUNTCURDEBIT: x.curDebit !== undefined ? x.curDebit : (x.debit || 0), AMOUNTCURCREDIT: x.curCredit !== undefined ? x.curCredit : (x.credit || 0),
      DEBITAMOUNT: '', CREDITAMOUNT: '', NOTE: x.note || '', REFERENCE: x.ref || ''
    }, 'add', ctx);
  });
  let dr = 0, cr = 0;
  lines.forEach(function (p) { dr += num_(p.rec.DEBITAMOUNT); cr += num_(p.rec.CREDITAMOUNT); });
  ph.rec.TOTALDEBIT = round_(dr, 3);
  ph.rec.TOTALCREDIT = round_(cr, 3);
  return { head: ph, lines: lines, t: t };
}

function writeVoucher_(ctx, v) {
  if (v.head.rec.VOUCHER === '@pending') {
    const no = nextSeq_('GLSEQ', 'GLPREFIX', 'JV-');
    v.head.rec.VOUCHER = no;
    v.lines.forEach(function (p) { p.rec.VOUCHER = no; });
  }
  commit_('glheaders', v.head, ctx);
  v.lines.forEach(function (p) { commit_('gllines', p, ctx); });
}

function balanced_(v) { return Math.abs(v.head.rec.TOTALDEBIT - v.head.rec.TOTALCREDIT) < 0.0005; }

/** قيد تلقائي مُرحّل (من يومية الأصول أو النقل). */
function autoVoucher_(ctx, head, legs) {
  head.SOURCE = head.SOURCE || 'FAJournal';
  const v = prepareVoucher_(ctx, head, legs, 'add');
  if (!balanced_(v)) throw new Error('القيد التلقائي غير متوازن');
  const now = stamp_();
  v.head.rec.SOURCE = head.SOURCE; v.head.rec.SOURCEREF = head.SOURCEREF || '';
  v.head.rec.STATUS = 'Posted'; v.head.rec.APPROVEDBY = ctx.s.username; v.head.rec.POSTEDBY = ctx.s.username; v.head.rec.POSTEDON = now;
  return v;
}

function newGLVoucherNumber(token) {
  const s = auth_(token);
  need_(s, 'gl', 'add');
  return withLock_(function () { return nextSeq_('GLSEQ', 'GLPREFIX', 'JV-'); });
}

/** حفظ قيد يدوي: الرأس + كل السطور (تُستبدل سطور القيد بالكامل). */
function saveGLVoucher(token, head, lines, isNew) {
  const s = auth_(token);
  need_(s, 'gl', isNew ? 'add' : 'edit');
  return withLock_(function () {
    const ctx = ctx_(s);
    head = head || {};
    if (!isNew) {
      const ex = ctx.find('glheaders', head.VOUCHER);
      if (!ex) throw new Error('القيد غير موجود');
      if (ex.STATUS !== 'Draft') throw new Error('لا يمكن تعديل قيد معتمد أو مُرحّل');
      if (ex.SOURCE && ex.SOURCE !== 'Manual') throw new Error('هذا قيد تلقائي ولا يمكن تعديله');
    }
    const legs = (lines || []).filter(function (x) {
      return x && (x.MAINACCOUNTID || x.PARTYID || num_(x.AMOUNTCURDEBIT) || num_(x.AMOUNTCURCREDIT));
    }).map(function (x) {
      return { pt: x.ACCOUNTTYPE && x.ACCOUNTTYPE !== 'Ledger' ? x.ACCOUNTTYPE : '', pid: x.PARTYID || '', acc: x.MAINACCOUNTID, s2: x.SUB2_OPERATINGUNIT, s1: x.SUB1_OPERATINGUNIT, ou: x.OPERATINGUNIT, cur: x.CURRENCYCODE, rate: toNum_(x.EXCHANGERATE, { ar: 'سعر الصرف' }),
        curDebit: toNum_(x.AMOUNTCURDEBIT, { ar: 'مدين' }) || 0, curCredit: toNum_(x.AMOUNTCURCREDIT, { ar: 'دائن' }) || 0, note: x.NOTE, ref: x.REFERENCE };
    });
    // التحقق أولاً (كل الأخطاء قبل أي كتابة)
    const gl = ctx.table('gllines');
    const old = [];
    gl.rows.forEach(function (r, i) { if (!isNew && low_(r.VOUCHER) === low_(head.VOUCHER)) old.push(i); });
    const h = Object.assign({}, head, { SOURCE: 'Manual' });
    // نحذف السطور القديمة من الذاكرة مؤقتاً حتى لا تتعارض المفاتيح، ثم نتحقق
    const v = prepareVoucher_(ctx, h, legs, isNew ? 'add' : 'edit');
    v.head.rec.SOURCE = 'Manual';
    if (!v.head.rec.STATUS) v.head.rec.STATUS = 'Draft';
    deleteRows_(gl, old);
    ctx.resetIndex('gllines');
    writeVoucher_(ctx, v);
    return { voucher: v.head.rec.VOUCHER, debit: v.head.rec.TOTALDEBIT, credit: v.head.rec.TOTALCREDIT };
  });
}

/** action: approve | unapprove | post | delete */
function glAction(token, voucher, action) {
  const s = auth_(token);
  need_(s, 'gl', action === 'post' ? 'post' : action === 'delete' ? 'delete' : 'approve');
  return withLock_(function () {
    const ctx = ctx_(s);
    const ht = ctx.table('glheaders');
    const h = ctx.find('glheaders', voucher);
    if (!h) throw new Error('القيد غير موجود');
    const hi = ht.rows.indexOf(h);
    const now = stamp_();
    if (action === 'delete') {
      if (h.STATUS !== 'Draft') throw new Error('يمكن حذف القيود في حالة مسودة فقط');
      const gl = ctx.table('gllines');
      const idx = [];
      gl.rows.forEach(function (r, i) { if (low_(r.VOUCHER) === low_(voucher)) idx.push(i); });
      deleteRows_(gl, idx);
      deleteRows_(ht, [hi]);
      return true;
    }
    if (h.SOURCE && h.SOURCE !== 'Manual') throw new Error('هذا قيد تلقائي مُرحّل');
    if (action === 'approve') {
      if (h.STATUS !== 'Draft') throw new Error('القيد ليس في حالة مسودة');
      if (Math.abs(num_(h.TOTALDEBIT) - num_(h.TOTALCREDIT)) >= 0.0005) throw new Error('القيد غير متوازن: المدين ' + h.TOTALDEBIT + ' ≠ الدائن ' + h.TOTALCREDIT);
      if (!num_(h.TOTALDEBIT)) throw new Error('القيد بدون مبالغ');
      h.STATUS = 'Approved'; h.APPROVEDBY = s.username;
    } else if (action === 'unapprove') {
      if (h.STATUS !== 'Approved') throw new Error('يمكن إلغاء الاعتماد للقيود المعتمدة فقط');
      h.STATUS = 'Draft'; h.APPROVEDBY = '';
    } else if (action === 'post') {
      if (h.STATUS !== 'Approved') throw new Error('يجب اعتماد القيد قبل الترحيل');
      checkOpenPeriod_(ctx, h.TRANSDATE, 'القيد');
      h.STATUS = 'Posted'; h.POSTEDBY = s.username; h.POSTEDON = now;
    } else {
      throw new Error('إجراء غير معروف');
    }
    writeRow_(ht, hi, h);
    return true;
  });
}

/* ============================== Fiscal years & closing ============================== */

/** يمنع أي حركة بتاريخ داخل سنة مالية مقفلة. */
function checkOpenPeriod_(ctx, date, what) {
  if (!date) return;
  const y = ctx.table('fiscal').rows.filter(function (r) {
    return r.YEAR && r.STATUS === 'Closed' && String(date) >= String(r.STARTDATE) && String(date) <= String(r.ENDDATE);
  })[0];
  if (y) throw new Error('لا يمكن حفظ أو ترحيل ' + (what || 'الحركة') + ' بتاريخ ' + date + ' لأن السنة المالية ' + y.YEAR + ' مقفلة');
}

/** ملخص قبل الإقفال: الإيرادات والمصروفات وصافي الربح + المستندات غير المُرحّلة. */
function fiscalSummary_(ctx, y) {
  const inYear = function (d) { return String(d) >= String(y.STARTDATE) && String(d) <= String(y.ENDDATE); };
  const heads = {};
  ctx.table('glheaders').rows.forEach(function (h) { if (h.VOUCHER) heads[low_(h.VOUCHER)] = h; });
  const accs = {};
  ctx.table('accounts').rows.forEach(function (a) { accs[low_(a.MAINACCOUNTID)] = a; });
  const groups = {};
  let revenue = 0, expense = 0, other = 0;
  ctx.table('gllines').rows.forEach(function (l) {
    const h = heads[low_(l.VOUCHER)];
    if (!h || h.STATUS !== 'Posted' || h.SOURCE === 'YearEnd' || !inYear(h.TRANSDATE)) return;
    const a = accs[low_(l.MAINACCOUNTID)];
    const t = a ? a.MAINACCOUNTTYPE : '';
    if (t !== 'Revenue' && t !== 'Expense' && t !== 'ProfitAndLoss') return;
    const net = num_(l.DEBITAMOUNT) - num_(l.CREDITAMOUNT);
    if (t === 'Revenue') revenue -= net; else if (t === 'Expense') expense += net; else other -= net;
    const k = [low_(l.MAINACCOUNTID), l.OPERATINGUNIT || '', l.SUB1_OPERATINGUNIT || '', l.SUB2_OPERATINGUNIT || ''].join('|');
    const g = groups[k] = groups[k] || { acc: l.MAINACCOUNTID, ou: l.OPERATINGUNIT || '', s1: l.SUB1_OPERATINGUNIT || '', s2: l.SUB2_OPERATINGUNIT || '', net: 0 };
    g.net += net;
  });
  const pending = {
    vouchers: ctx.table('glheaders').rows.filter(function (h) { return h.VOUCHER && h.STATUS !== 'Posted' && inYear(h.TRANSDATE); }).length,
    faLines: ctx.table('journal').rows.filter(function (j) { return j.LINEID && j.STATUS !== 'Posted' && inYear(j.TRANSDATE); }).length,
    transfers: ctx.table('transfers').rows.filter(function (j) { return j.LINEID && j.STATUS !== 'Posted' && inYear(j.TRANSFERDATE); }).length
  };
  return { revenue: round_(revenue, 3), expense: round_(expense, 3), other: round_(other, 3), net: round_(revenue - expense + other, 3),
    groups: Object.keys(groups).map(function (k) { return groups[k]; }).filter(function (g) { return Math.abs(g.net) >= 0.0005; }), pending: pending };
}

function fiscalPreview(token, year) {
  const s = auth_(token);
  need_(s, 'fiscal', 'view');
  const ctx = ctx_(s);
  const y = ctx.find('fiscal', year);
  if (!y) throw new Error('السنة المالية غير موجودة');
  const sum = fiscalSummary_(ctx, y);
  return { year: y.YEAR, start: y.STARTDATE, end: y.ENDDATE, status: y.STATUS, revenue: sum.revenue, expense: sum.expense, other: sum.other, net: sum.net,
    lines: sum.groups.length, pending: sum.pending, account: y.RETAINEDEARNINGSACCOUNT || getSettings_().RETAINEDEARNINGSACCOUNT || '' };
}

/** إقفال السنة: ترحيل أرصدة الإيرادات والمصروفات (لكل وحدة) إلى حساب الأرباح المحتجزة ثم قفل الفترة. */
function fiscalClose(token, year, reAccount) {
  const s = auth_(token);
  need_(s, 'fiscal', 'post');
  return withLock_(function () {
    const ctx = ctx_(s);
    const t = ctx.table('fiscal');
    const y = ctx.find('fiscal', year);
    if (!y) throw new Error('السنة المالية غير موجودة');
    if (y.STATUS === 'Closed') throw new Error('السنة مقفلة مسبقاً');
    const earlier = t.rows.filter(function (r) { return r.YEAR && r.STATUS !== 'Closed' && String(r.ENDDATE) < String(y.STARTDATE); })[0];
    if (earlier) throw new Error('أقفل السنة السابقة ' + earlier.YEAR + ' أولاً');
    const sum = fiscalSummary_(ctx, y);
    const p = sum.pending;
    if (p.vouchers || p.faLines || p.transfers) {
      throw new Error('توجد مستندات غير مُرحّلة في السنة: ' + p.vouchers + ' قيد، ' + p.faLines + ' سطر يومية أصول، ' + p.transfers + ' سطر نقل — رحّلها أو احذفها قبل الإقفال');
    }
    const re = postingAccount_(ctx, reAccount || y.RETAINEDEARNINGSACCOUNT || getSettings_().RETAINEDEARNINGSACCOUNT, 'حساب الأرباح المحتجزة');
    if (!re) throw new Error('حدد حساب الأرباح المحتجزة');
    const reAcc = ctx.find('accounts', re);
    if (reAcc.MAINACCOUNTTYPE !== 'Equity') throw new Error('حساب الأرباح المحتجزة يجب أن يكون من نوع حقوق ملكية');
    let v = null;
    if (sum.groups.length) {
      const legs = [], perUnit = {};
      sum.groups.forEach(function (g) {
        legs.push({ acc: g.acc, ou: g.ou, s1: g.s1, s2: g.s2, debit: g.net < 0 ? round_(-g.net, 3) : 0, credit: g.net > 0 ? round_(g.net, 3) : 0,
          note: 'إقفال ' + g.acc + ' — سنة ' + y.YEAR });
        const k = [g.ou, g.s1, g.s2].join('|');
        const u = perUnit[k] = perUnit[k] || { ou: g.ou, s1: g.s1, s2: g.s2, net: 0 };
        u.net += g.net;
      });
      Object.keys(perUnit).forEach(function (k) {
        const u = perUnit[k], n = round_(u.net, 3);
        if (!n) return;
        legs.push({ acc: re, ou: u.ou, s1: u.s1, s2: u.s2, debit: n > 0 ? n : 0, credit: n < 0 ? -n : 0,
          note: (n < 0 ? 'صافي ربح' : 'صافي خسارة') + ' سنة ' + y.YEAR });
      });
      v = autoVoucher_(ctx, { TRANSDATE: y.ENDDATE, JOURNALNAME: getSettings_().CLOSINGJOURNALNAME || 'YE',
        DESCRIPTION: 'قيد إقفال السنة المالية ' + y.YEAR, REFERENCE: 'FY' + y.YEAR, SOURCE: 'YearEnd', SOURCEREF: y.YEAR }, legs);
      writeVoucher_(ctx, v);
    }
    y.STATUS = 'Closed';
    y.RETAINEDEARNINGSACCOUNT = re;
    y.CLOSINGVOUCHER = v ? v.head.rec.VOUCHER : '';
    y.CLOSEDBY = s.username; y.CLOSEDON = stamp_();
    writeRow_(t, t.rows.indexOf(y), y);
    return { year: y.YEAR, voucher: y.CLOSINGVOUCHER, net: sum.net };
  });
}

/** إعادة فتح سنة مقفلة: حذف قيد الإقفال وفتح الفترة. */
function fiscalReopen(token, year) {
  const s = auth_(token);
  need_(s, 'fiscal', 'post');
  return withLock_(function () {
    const ctx = ctx_(s);
    const t = ctx.table('fiscal');
    const y = ctx.find('fiscal', year);
    if (!y) throw new Error('السنة المالية غير موجودة');
    if (y.STATUS !== 'Closed') throw new Error('السنة ليست مقفلة');
    const later = t.rows.filter(function (r) { return r.YEAR && r.STATUS === 'Closed' && String(r.STARTDATE) > String(y.ENDDATE); })[0];
    if (later) throw new Error('أعد فتح السنة اللاحقة ' + later.YEAR + ' أولاً');
    if (y.CLOSINGVOUCHER) {
      const gl = ctx.table('gllines'), ht = ctx.table('glheaders');
      const idx = [];
      gl.rows.forEach(function (r, i) { if (low_(r.VOUCHER) === low_(y.CLOSINGVOUCHER)) idx.push(i); });
      deleteRows_(gl, idx);
      const hi = ht.rows.findIndex(function (h) { return low_(h.VOUCHER) === low_(y.CLOSINGVOUCHER); });
      if (hi >= 0) deleteRows_(ht, [hi]);
    }
    y.STATUS = 'Open'; y.CLOSINGVOUCHER = ''; y.CLOSEDBY = ''; y.CLOSEDON = '';
    writeRow_(t, t.rows.indexOf(y), y);
    return true;
  });
}

/* ============================== Sales & purchase invoices ============================== */

function invKind_(type) { const k = INV_KIND[type]; if (!k) throw new Error('نوع غير معروف: ' + type); return k; }
function invParty_(type) { return PARTY[(INV_KIND[type] || INV_KIND.Sales).side === 'Purchase' ? 'Vendor' : 'Customer']; }
/** المتبقي على الفاتورة = الإجمالي − المسدد − المخصوم بإشعارات دائنة. */
function invRem_(h) { return round_(num_(h.TOTALAMOUNT) - num_(h.PAIDAMOUNT) - num_(h.CREDITAMOUNT), 3); }

/** قيد الفاتورة: المبيعات = من حـ/ العميل إلى حـ/ الإيرادات + الضريبة، والمشتريات عكسها. */
function invoiceLegs_(ctx, h, lines, reverse) {
  const st = getSettings_();
  const K = invKind_(h.INVOICETYPE);
  const sales = K.side === 'Sales';
  const cur = h.CURRENCYCODE, rate = num_(h.EXCHANGERATE) || 1;
  // المبيعات: العميل مدين، والإشعار الدائن (مرتجع) يعكس ذلك، والإلغاء يعكسه مرة أخرى
  const partyDr = (sales !== !!K.credit) !== !!reverse;
  const note = K.ar + ' ' + h.INVOICEID + (h.ORIGINVOICE ? ' (للفاتورة ' + h.ORIGINVOICE + ')' : '');
  const legs = [{ pt: sales ? 'Customer' : 'Vendor', pid: h.PARTYID, cur: cur, rate: rate, curDebit: partyDr ? num_(h.TOTALAMOUNT) : 0,
    curCredit: partyDr ? 0 : num_(h.TOTALAMOUNT), note: h.DESCRIPTION || note, ref: h.PARTYREF || h.INVOICEID, s2: h.SUB2_OPERATINGUNIT || '' }];
  lines.forEach(function (l) {
    const a = num_(l.LINEAMOUNT);
    if (!a) return;
    legs.push({ acc: l.MAINACCOUNTID, s2: l.SUB2_OPERATINGUNIT, s1: l.SUB1_OPERATINGUNIT, ou: l.OPERATINGUNIT, cur: cur, rate: rate,
      curDebit: partyDr ? 0 : a, curCredit: partyDr ? a : 0, note: l.ITEMDESCRIPTION, ref: h.INVOICEID });
  });
  const tax = num_(h.TAXAMOUNT);
  if (tax) {
    const acc = sales ? st.SALESTAXACCOUNT : st.PURCHTAXACCOUNT;
    if (!acc) throw new Error('حدد حساب ضريبة ' + (sales ? 'المبيعات' : 'المشتريات') + ' من الإعدادات');
    legs.push({ acc: acc, s2: h.SUB2_OPERATINGUNIT || '', cur: cur, rate: rate, curDebit: partyDr ? 0 : tax, curCredit: partyDr ? tax : 0, note: 'ضريبة ' + note, ref: h.INVOICEID });
  }
  return legs;
}

function postInvoice_(ctx, h, lines, reverse, date) {
  const K = invKind_(h.INVOICETYPE);
  const v = autoVoucher_(ctx, { TRANSDATE: date || h.INVOICEDATE, JOURNALNAME: K.jn,
    DESCRIPTION: (reverse ? 'إلغاء ' : '') + K.ar + ' ' + h.INVOICEID + (h.ORIGINVOICE ? ' (للفاتورة ' + h.ORIGINVOICE + ')' : '') + ' — ' + (h.PARTYNAME || h.PARTYID),
    REFERENCE: h.PARTYREF || h.INVOICEID, SOURCE: 'Invoice', SOURCEREF: h.INVOICEID }, invoiceLegs_(ctx, h, lines, reverse));
  writeVoucher_(ctx, v);
  return v.head.rec.VOUCHER;
}

function invLines_(ctx, id) {
  return ctx.table('invlines').rows.filter(function (l) { return low_(l.INVOICEID) === low_(id); })
    .sort(function (a, b) { return num_(a.LINENUM) - num_(b.LINENUM); });
}

/** ربط الإشعار بالفاتورة الأصلية: نفس الطرف والجانب، والإشعار الدائن لا يتجاوز المتبقي. يرجع المبلغ بعملة الفاتورة الأصلية. */
function checkNoteLink_(ctx, h) {
  const K = invKind_(h.INVOICETYPE);
  if (!h.ORIGINVOICE) {
    if (K.credit) throw new Error('حدد الفاتورة الأصلية التي يخصم منها الإشعار');
    return 0;
  }
  const o = ctx.find('invoices', h.ORIGINVOICE);
  if (!o) throw new Error('الفاتورة الأصلية غير موجودة "' + h.ORIGINVOICE + '"');
  h.ORIGINVOICE = o.INVOICEID;
  const OK = INV_KIND[o.INVOICETYPE] || {};
  if (OK.side !== K.side || OK.credit) throw new Error('الفاتورة ' + o.INVOICEID + ' لا تصلح أصلاً لهذا الإشعار');
  if (low_(o.PARTYID) !== low_(h.PARTYID)) throw new Error('الفاتورة ' + o.INVOICEID + ' تخص ' + o.PARTYID + ' وليس ' + h.PARTYID);
  if (o.STATUS !== 'Posted') throw new Error('الفاتورة الأصلية ' + o.INVOICEID + ' غير مُرحّلة');
  if (!K.credit) return 0;
  const applied = round_(num_(h.TOTALAMOUNT) * (num_(h.EXCHANGERATE) || 1) / (num_(o.EXCHANGERATE) || 1), 3);
  // لا يتجاوز الإشعار ما لم يُخصم بعد من إجمالي الفاتورة؛ إن تجاوز المتبقي (فاتورة مسددة) يصبح الفرق رصيداً دائناً للطرف
  const cap = round_(num_(o.TOTALAMOUNT) - num_(o.CREDITAMOUNT), 3);
  if (applied > cap + 0.0005) throw new Error('الإشعار (' + applied + ') أكبر من قيمة الفاتورة ' + o.INVOICEID + ' غير المخصومة (' + cap + ' ' + o.CURRENCYCODE + ')');
  return applied;
}

/** إشعار دائن مُرحّل: يُخصم من الفاتورة الأصلية ويُعتبر الإشعار نفسه مُسوّى. sign = -1 عند الإلغاء. */
function applyNote_(ctx, h, sign) {
  if (!invKind_(h.INVOICETYPE).credit || !h.ORIGINVOICE) return;
  const o = ctx.find('invoices', h.ORIGINVOICE);
  if (!o) return;
  const applied = round_(num_(h.TOTALAMOUNT) * (num_(h.EXCHANGERATE) || 1) / (num_(o.EXCHANGERATE) || 1), 3);
  o.CREDITAMOUNT = Math.max(0, round_(num_(o.CREDITAMOUNT) + sign * applied, 3));
  ctx.touch('invoices', o);
  h.PAIDAMOUNT = sign > 0 ? num_(h.TOTALAMOUNT) : 0;
}

/** حفظ فاتورة (الرأس + السطور) وترحيلها مباشرة إذا post = true. */
function saveInvoice(token, head, lines, isNew, post) {
  const s = auth_(token);
  need_(s, 'invoices', isNew ? 'add' : 'edit');
  if (post) need_(s, 'invoices', 'post');
  return withLock_(function () {
    const ctx = ctx_(s);
    head = Object.assign({}, head || {});
    if (isNew) head.INVOICEID = '@pending';
    else {
      const ex = ctx.find('invoices', head.INVOICEID);
      if (!ex) throw new Error('الفاتورة غير موجودة');
      if (ex.STATUS !== 'Draft') throw new Error('لا يمكن تعديل فاتورة مُرحّلة أو ملغاة');
    }
    const ph = prepare_('invoices', head, isNew ? 'add' : 'edit', ctx);
    ph.rec.STATUS = 'Draft';
    ph.rec.PAIDAMOUNT = 0; ph.rec.CREDITAMOUNT = 0;
    const id = ph.rec.INVOICEID;
    const K = invKind_(ph.rec.INVOICETYPE);
    ctx.invType = K.side;
    ctx.invCredit = !!K.credit;
    const rows = (lines || []).filter(function (x) { return x && (x.ITEMID || x.ITEMDESCRIPTION || x.MAINACCOUNTID || num_(x.UNITPRICE)); });
    if (!rows.length) throw new Error((ph.rec.INVOICETYPE === K.side ? 'الفاتورة' : 'الإشعار') + ' لا يحتوي على سطور');
    let sub = 0, tax = 0;
    const pl = rows.map(function (x, i) {
      const p = prepare_('invlines', {
        INVOICEID: id, LINENUM: i + 1, ITEMID: x.ITEMID || '', ITEMDESCRIPTION: x.ITEMDESCRIPTION || '', UNIT: x.UNIT || '', MAINACCOUNTID: x.MAINACCOUNTID || '',
        SUB2_OPERATINGUNIT: ph.rec.SUB2_OPERATINGUNIT || '', QTY: toNum_(x.QTY === '' || x.QTY === undefined ? 1 : x.QTY, { ar: 'سطر ' + (i + 1) + ' — الكمية' }),
        UNITPRICE: toNum_(x.UNITPRICE, { ar: 'سطر ' + (i + 1) + ' — السعر' }), TAXPERCENT: toNum_(x.TAXPERCENT || 0, { ar: 'سطر ' + (i + 1) + ' — الضريبة' })
      }, 'add', ctx);
      sub += num_(p.rec.LINEAMOUNT); tax += num_(p.rec.TAXAMOUNT);
      return p;
    });
    ph.rec.SUBTOTAL = round_(sub, 3); ph.rec.TAXAMOUNT = round_(tax, 3); ph.rec.TOTALAMOUNT = round_(sub + tax, 3);
    if (!(ph.rec.TOTALAMOUNT > 0)) throw new Error('الإجمالي يجب أن يكون أكبر من صفر');
    checkNoteLink_(ctx, ph.rec);
    // التحقق من القيد قبل أي كتابة
    if (post) prepareVoucher_(ctx, { TRANSDATE: ph.rec.INVOICEDATE, JOURNALNAME: K.jn },
      invoiceLegs_(ctx, ph.rec, pl.map(function (p) { return p.rec; })), 'add');
    const lt = ctx.table('invlines');
    const old = [];
    lt.rows.forEach(function (r, i) { if (!isNew && low_(r.INVOICEID) === low_(id)) old.push(i); });
    deleteRows_(lt, old);
    ctx.resetIndex('invlines');
    if (id === '@pending') {
      const no = nextSeq_(K.seq[0], K.seq[1], K.seq[2]);
      ph.rec.INVOICEID = no;
      pl.forEach(function (p) { p.rec.INVOICEID = no; });
    }
    if (post) {
      ph.rec.GLVOUCHER = postInvoice_(ctx, ph.rec, pl.map(function (p) { return p.rec; }));
      ph.rec.STATUS = 'Posted'; ph.rec.POSTEDBY = s.username; ph.rec.POSTEDON = stamp_();
      applyNote_(ctx, ph.rec, 1);
    }
    commit_('invoices', ph, ctx);
    pl.forEach(function (p) { commit_('invlines', p, ctx); });
    ctx.flush();
    return { id: ph.rec.INVOICEID, status: ph.rec.STATUS, total: ph.rec.TOTALAMOUNT, gl: ph.rec.GLVOUCHER || '' };
  });
}

/** action: post | cancel (قيد عكسي) | delete (مسودة فقط) */
function invoiceAction(token, id, action, date) {
  const s = auth_(token);
  need_(s, 'invoices', action === 'delete' ? 'delete' : 'post');
  return withLock_(function () {
    const ctx = ctx_(s);
    const t = ctx.table('invoices');
    const h = ctx.find('invoices', id);
    if (!h) throw new Error('الفاتورة غير موجودة');
    const i = t.rows.indexOf(h);
    if (action === 'delete') {
      if (h.STATUS !== 'Draft') throw new Error('يمكن حذف الفواتير في حالة مسودة فقط — الفاتورة المُرحّلة تُلغى');
      const lt = ctx.table('invlines');
      const idx = [];
      lt.rows.forEach(function (r, k) { if (low_(r.INVOICEID) === low_(id)) idx.push(k); });
      deleteRows_(lt, idx);
      deleteRows_(t, [i]);
      return true;
    }
    if (action === 'post') {
      if (h.STATUS !== 'Draft') throw new Error('الفاتورة مُرحّلة مسبقاً');
      checkOpenPeriod_(ctx, h.INVOICEDATE, 'الفاتورة');
      checkNoteLink_(ctx, h);
      h.GLVOUCHER = postInvoice_(ctx, h, invLines_(ctx, id));
      h.STATUS = 'Posted'; h.POSTEDBY = s.username; h.POSTEDON = stamp_();
      applyNote_(ctx, h, 1);
    } else if (action === 'cancel') {
      if (h.STATUS !== 'Posted') throw new Error('يمكن إلغاء الفواتير المُرحّلة فقط');
      const K = invKind_(h.INVOICETYPE);
      if (!K.credit && num_(h.PAIDAMOUNT) > 0.0005) throw new Error('الفاتورة عليها سداد ' + h.PAIDAMOUNT + ' — ألغِ سندات القبض/الصرف المرتبطة أولاً');
      if (num_(h.CREDITAMOUNT) > 0.0005) throw new Error('على الفاتورة إشعارات دائنة بمبلغ ' + h.CREDITAMOUNT + ' — ألغِها أولاً');
      const d = date ? toDate_(date, ctx.dateFmt, { ar: 'تاريخ الإلغاء' }) : today_();
      if (d < String(h.INVOICEDATE)) throw new Error('تاريخ الإلغاء قبل تاريخ الفاتورة');
      h.REVERSALVOUCHER = postInvoice_(ctx, h, invLines_(ctx, id), true, d);
      h.STATUS = 'Cancelled';
      applyNote_(ctx, h, -1);
    } else throw new Error('إجراء غير معروف');
    writeRow_(t, i, h);
    ctx.flush();
    return { gl: action === 'post' ? h.GLVOUCHER : h.REVERSALVOUCHER };
  });
}

/** يضيف (أو يطرح عند الإلغاء) مبلغ السند على المسدد من الفاتورة المرتبطة به. */
function applyInvoicePay_(ctx, r, sign) {
  if (!r.INVOICEID) return;
  const h = ctx.find('invoices', r.INVOICEID);
  if (!h) throw new Error('الفاتورة غير موجودة "' + r.INVOICEID + '"');
  const applied = round_(num_(r.AMOUNT) * (num_(r.EXCHANGERATE) || 1) / (num_(h.EXCHANGERATE) || 1), 3);
  h.PAIDAMOUNT = round_(num_(h.PAIDAMOUNT) + sign * applied, 3);
  if (h.PAIDAMOUNT < 0) h.PAIDAMOUNT = 0;
  ctx.touch('invoices', h);
}

/** التحقق من ربط السند بفاتورة: نفس الطرف، نوع مطابق، ولا يتجاوز المتبقي. */
function checkInvoiceLink_(ctx, rec) {
  if (!rec.INVOICEID) return;
  const h = ctx.find('invoices', rec.INVOICEID);
  if (!h) throw new Error('الفاتورة غير موجودة "' + rec.INVOICEID + '"');
  rec.INVOICEID = h.INVOICEID;
  if (h.STATUS !== 'Posted') throw new Error('الفاتورة ' + h.INVOICEID + ' غير مُرحّلة');
  const K = invKind_(h.INVOICETYPE);
  if (K.credit) throw new Error(h.INVOICEID + ' إشعار دائن مُسوّى بالخصم من فاتورته ولا يُربط بسند');
  const want = K.side === 'Sales' ? ['Receipt', 'Customer'] : ['Payment', 'Vendor'];
  if (rec.VOUCHERTYPE !== want[0] || rec.PARTYTYPE !== want[1]) {
    throw new Error(h.INVOICEID + (K.side === 'Sales' ? ' مستند مبيعات ويُحصَّل بسند قبض من العميل' : ' مستند مشتريات ويُسدَّد بسند صرف للمورد'));
  }
  if (low_(h.PARTYID) !== low_(rec.PARTYID)) throw new Error('الفاتورة ' + h.INVOICEID + ' تخص ' + h.PARTYID + ' وليس ' + rec.PARTYID);
  const applied = num_(rec.AMOUNT) * (num_(rec.EXCHANGERATE) || 1) / (num_(h.EXCHANGERATE) || 1);
  const rem = invRem_(h);
  if (applied > rem + 0.0005) throw new Error('المبلغ أكبر من المتبقي على الفاتورة ' + h.INVOICEID + ' (' + rem + ' ' + h.CURRENCYCODE + ')');
}

/* ============================== Aging (receivables / payables) ============================== */

/** فواتير الطرف المُرحّلة مع المسدد منها حتى تاريخ معيّن (لتقرير أعمار الديون). */
function agingData(token, type, asOf) {
  const s = auth_(token);
  const P = PARTY[type];
  if (!P || type === 'Bank') throw new Error('نوع غير معروف');
  need_(s, P.perm, 'view');
  asOf = String(asOf || today_());
  const side = type === 'Customer' ? 'Sales' : 'Purchase';
  const paid = {};
  readTable_('CashVouchers').rows.forEach(function (c) {
    if (!c.INVOICEID || c.STATUS !== 'Posted' || String(c.TRANSDATE) > asOf) return;
    const k = low_(c.INVOICEID);
    paid[k] = (paid[k] || 0) + num_(c.AMOUNT) * (num_(c.EXCHANGERATE) || 1);   // بالعملة المحلية
  });
  const all = readTable_('Invoices').rows;
  // الإشعارات الدائنة المُرحّلة حتى التاريخ تُخصم من فواتيرها
  all.forEach(function (n) {
    if (!n.ORIGINVOICE || n.STATUS !== 'Posted' || !(INV_KIND[n.INVOICETYPE] || {}).credit || String(n.INVOICEDATE) > asOf) return;
    const k = low_(n.ORIGINVOICE);
    paid[k] = (paid[k] || 0) + num_(n.TOTALAMOUNT) * (num_(n.EXCHANGERATE) || 1);
  });
  const invoices = all.filter(function (h) {
    const K = INV_KIND[h.INVOICETYPE] || {};
    return h.INVOICEID && K.side === side && !K.credit && h.STATUS === 'Posted' && String(h.INVOICEDATE) <= asOf;
  }).map(function (h) {
    const rate = num_(h.EXCHANGERATE) || 1;
    const p = round_((paid[low_(h.INVOICEID)] || 0) / rate, 3);
    return { id: String(h.INVOICEID), type: h.INVOICETYPE, party: String(h.PARTYID), name: String(h.PARTYNAME || ''), date: String(h.INVOICEDATE), due: String(h.DUEDATE || h.INVOICEDATE),
      cur: String(h.CURRENCYCODE || ''), rate: rate, total: num_(h.TOTALAMOUNT), paid: p, rem: round_(num_(h.TOTALAMOUNT) - p, 3), ref: String(h.PARTYREF || '') };
  }).filter(function (x) { return x.rem > 0.0005; });
  return { asOf: asOf, invoices: invoices };
}

/* ============================== Fixed assets opening balances ============================== */

/** الأرصدة الافتتاحية للأصول = قيم الدفتر الأساسي غير المُرحّلة عبر يومية الأصول (مثل الأصول المستوردة من D365). */
function faOpeningRows_(ctx) {
  const posted = {};
  ctx.table('journal').rows.forEach(function (j) {
    if (j.STATUS !== 'Posted' || !j.FIXEDASSETNUMBER) return;
    const k = low_(j.FIXEDASSETNUMBER) + '|' + low_(j.BOOKID || 'General');
    const p = posted[k] = posted[k] || { cost: 0, accum: 0 };
    const amt = num_(j.DEBITAMOUNT) - num_(j.CREDITAMOUNT);
    if (j.TRANSACTIONTYPE === 'Acquisition' || j.TRANSACTIONTYPE === 'AcquisitionAdj') p.cost += amt;
    else if (DEP_TYPES[j.TRANSACTIONTYPE]) p.accum -= amt;
  });
  const glOk = function (v) { return v && ctx.find('glheaders', v); };
  const rows = [], errors = [];
  let done = 0;
  const prim = primaryBooks_(ctx.table('books').rows);
  Object.keys(prim).forEach(function (k) {
    const b = prim[k];
    if (b.STATUS !== 'Open' && b.STATUS !== 'Suspended') return;
    if (glOk(b.OPENINGVOUCHER)) { done++; return; }
    const p = posted[low_(b.FIXEDASSETNUMBER) + '|' + low_(b.BOOKID || 'General')] || { cost: 0, accum: 0 };
    const cost = round_(num_(b.ACQUISITIONPRICE) - p.cost, 3), accum = round_(num_(b.ACCUMULATEDDEPRECIATION) - p.accum, 3);
    if (Math.abs(cost) < 0.0005 && Math.abs(accum) < 0.0005) return;
    const a = ctx.asset(b.FIXEDASSETNUMBER) || {};
    const g = ctx.find('groups', a.FIXEDASSETGROUPID) || {};
    const r = { asset: b.FIXEDASSETNUMBER, book: b.BOOKID, name: a.NAME || '', group: a.FIXEDASSETGROUPID || '', ou: a.OPERATINGUNIT || '',
      s1: a.SUB1_OPERATINGUNIT || '', s2: a.SUB2_OPERATINGUNIT || '', cost: cost, accum: accum, nbv: round_(cost - accum, 3),
      assetAcc: g.ASSETACCOUNT || '', accumAcc: g.ACCUMDEPACCOUNT || '' };
    if (cost && !r.assetAcc) errors.push(r.asset + ': مجموعة الأصل ' + (r.group || '—') + ' بدون حساب أصل');
    if (accum && !r.accumAcc) errors.push(r.asset + ': مجموعة الأصل ' + (r.group || '—') + ' بدون حساب مجمع إهلاك');
    rows.push(r);
  });
  rows.sort(function (x, y) { return String(x.asset).localeCompare(String(y.asset), undefined, { numeric: true }); });
  return { rows: rows, errors: errors, done: done };
}

function faOpeningData(token) {
  const s = auth_(token);
  need_(s, 'books', 'view');
  const ctx = ctx_(s);
  const r = faOpeningRows_(ctx);
  r.offset = getSettings_().OPENINGBALANCEACCOUNT || '';
  return r;
}

/** ينشئ قيد الأرصدة الافتتاحية للأصول (مسودة): من حـ/ الأصول ← إلى حـ/ مجمع الإهلاك + حـ/ الأرصدة الافتتاحية، مجمّعاً حسب الحساب والوحدة. */
function createFAOpeningVoucher(token, date, offsetAccount) {
  const s = auth_(token);
  need_(s, 'books', 'edit');
  need_(s, 'gl', 'add');
  return withLock_(function () {
    const ctx = ctx_(s);
    const d = faOpeningRows_(ctx);
    if (!d.rows.length) throw new Error('لا توجد أرصدة افتتاحية للأصول غير مُرحّلة');
    if (d.errors.length) throw new Error('أكمل حسابات مجموعات الأصول أولاً:\n' + d.errors.slice(0, 10).join('\n'));
    const off = postingAccount_(ctx, offsetAccount || getSettings_().OPENINGBALANCEACCOUNT, 'حساب الأرصدة الافتتاحية');
    if (!off) throw new Error('حدد حساب الأرصدة الافتتاحية (المقابل)');
    date = toDate_(date || today_(), ctx.dateFmt, { ar: 'التاريخ' });
    const local = String(getSettings_().CURRENCY || '').toUpperCase();
    const agg = {}, order = [];
    const add = function (acc, r, amt, what) {
      if (Math.abs(amt) < 0.0005) return;
      const k = [low_(acc), low_(r.ou), low_(r.s1), low_(r.s2)].join('|');
      if (!agg[k]) { agg[k] = { acc: acc, ou: r.ou, s1: r.s1, s2: r.s2, amt: 0, n: 0, what: what }; order.push(k); }
      agg[k].amt += amt; agg[k].n++;
    };
    let net = 0;
    d.rows.forEach(function (r) { add(r.assetAcc, r, r.cost, 'تكلفة'); add(r.accumAcc, r, -r.accum, 'مجمع إهلاك'); net += r.cost - r.accum; });
    const legs = order.map(function (k) {
      const x = agg[k], amt = round_(x.amt, 3);
      return { acc: x.acc, ou: x.ou, s1: x.s1, s2: x.s2, cur: local, rate: 1, curDebit: amt > 0 ? amt : 0, curCredit: amt < 0 ? -amt : 0,
        note: 'رصيد افتتاحي — ' + x.what + ' الأصول (' + x.n + ' أصل)', ref: 'FA-OPENING' };
    }).filter(function (l) { return l.curDebit || l.curCredit; });
    net = round_(net, 3);
    if (net) legs.push({ acc: off, cur: local, rate: 1, curDebit: net < 0 ? -net : 0, curCredit: net > 0 ? net : 0, note: 'مقابل الأرصدة الافتتاحية للأصول الثابتة', ref: 'FA-OPENING' });
    const v = prepareVoucher_(ctx, { TRANSDATE: date, JOURNALNAME: 'GJ', DESCRIPTION: 'الأرصدة الافتتاحية للأصول الثابتة (' + d.rows.length + ' أصل)', REFERENCE: 'FA-OPENING' }, legs, 'add');
    v.head.rec.SOURCE = 'Manual';
    v.head.rec.STATUS = 'Draft';
    if (!balanced_(v)) throw new Error('القيد غير متوازن');
    writeVoucher_(ctx, v);
    const bt = ctx.table('books'), bi = ctx.index('books');
    d.rows.forEach(function (r) {
      const i = bi[low_(r.asset) + '|' + low_(r.book || 'General')];
      if (i === undefined) return;
      const b = bt.rows[i];
      b.OPENINGVOUCHER = v.head.rec.VOUCHER; b.OPENINGCOST = r.cost; b.OPENINGACCUM = r.accum;
      writeRow_(bt, i, b);
    });
    return { voucher: v.head.rec.VOUCHER, count: d.rows.length, cost: round_(d.rows.reduce(function (a, r) { return a + r.cost; }, 0), 3),
      accum: round_(d.rows.reduce(function (a, r) { return a + r.accum; }, 0), 3) };
  });
}

/* ============================== Sales & purchases dashboard ============================== */

/** بيانات لوحة المبيعات والمشتريات لسنة: شهرياً، أكبر العملاء والموردين، المستحقات والمتأخرات، والمرتجعات (بالعملة المحلية، قبل الضريبة). */
function salesDashboard(token, year) {
  const s = auth_(token);
  need_(s, 'invoices', 'view');
  year = String(year || new Date().getFullYear());
  const today = today_();
  const z = function () { return [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; };
  const out = { year: year, sales: z(), purch: z(), salesRet: z(), purchRet: z(), customers: {}, vendors: {},
    ar: { total: 0, overdue: 0, b4: 0, n: 0 }, ap: { total: 0, overdue: 0, b4: 0, n: 0 }, counts: { si: 0, pi: 0, notes: 0 }, overdueList: [] };
  readTable_('Invoices').rows.forEach(function (h) {
    if (!h.INVOICEID || h.STATUS !== 'Posted') return;
    const K = INV_KIND[h.INVOICETYPE];
    if (!K) return;
    const rate = num_(h.EXCHANGERATE) || 1, net = num_(h.SUBTOTAL) * rate, sales = K.side === 'Sales';
    // المستحقات والمتأخرات (كل السنوات، حتى اليوم)
    if (!K.credit) {
      const rem = invRem_(h) * rate;
      if (rem > 0.0005) {
        const box = sales ? out.ar : out.ap;
        const due = String(h.DUEDATE || h.INVOICEDATE);
        box.total += rem; box.n++;
        if (due < today) {
          box.overdue += rem;
          const days = Math.round((Date.parse(today + 'T00:00:00Z') - Date.parse(due + 'T00:00:00Z')) / 86400000);
          if (days > 90) box.b4 += rem;
          out.overdueList.push({ id: h.INVOICEID, type: h.INVOICETYPE, side: K.side, party: h.PARTYID, name: h.PARTYNAME || '', due: due, days: days, rem: round_(rem, 3) });
        }
      }
    }
    if (String(h.INVOICEDATE).slice(0, 4) !== year) return;
    const m = Number(String(h.INVOICEDATE).slice(5, 7)) - 1;
    if (!(m >= 0 && m < 12)) return;
    if (K.credit) { (sales ? out.salesRet : out.purchRet)[m] += net; out.counts.notes++; }
    else { (sales ? out.sales : out.purch)[m] += net; if (h.INVOICETYPE === 'Sales') out.counts.si++; else if (h.INVOICETYPE === 'Purchase') out.counts.pi++; else out.counts.notes++; }
    const map = sales ? out.customers : out.vendors, pk = low_(h.PARTYID);
    map[pk] = map[pk] || { c: h.PARTYID, n: h.PARTYNAME || '', v: 0 };
    map[pk].v += K.credit ? -net : net;
  });
  const top = function (m) { return Object.keys(m).map(function (k) { m[k].v = round_(m[k].v, 3); return m[k]; }).sort(function (a, b) { return b.v - a.v; }).slice(0, 8); };
  out.customers = top(out.customers); out.vendors = top(out.vendors);
  ['sales', 'purch', 'salesRet', 'purchRet'].forEach(function (k) { out[k] = out[k].map(function (x) { return round_(x, 3); }); });
  ['ar', 'ap'].forEach(function (k) { ['total', 'overdue', 'b4'].forEach(function (f) { out[k][f] = round_(out[k][f], 3); }); });
  out.overdueList.sort(function (a, b) { return b.days - a.days; });
  out.overdueList = out.overdueList.slice(0, 10);
  return out;
}

/* ============================== Budgets: header + lines ============================== */

/** حفظ موازنة كاملة (الرأس + كل السطور؛ تُستبدل السطور بالكامل). */
function saveBudget(token, head, lines, isNew) {
  const s = auth_(token);
  need_(s, 'budget', isNew ? 'add' : 'edit');
  return withLock_(function () {
    const ctx = ctx_(s);
    head = head || {};
    if (!isNew) {
      const ex = ctx.find('budgetheads', head.BUDGETID);
      if (!ex) throw new Error('الموازنة غير موجودة');
      if (ex.STATUS !== 'Draft') throw new Error('لا يمكن تعديل موازنة معتمدة أو مُرحّلة');
    } else head.BUDGETID = '@pending';
    const ph = prepare_('budgetheads', head, isNew ? 'add' : 'edit', ctx);
    ph.rec.STATUS = 'Draft';
    const id = ph.rec.BUDGETID;
    const rows = (lines || []).filter(function (x) { return x && (x.MAINACCOUNTID || num_(x.AMOUNT)); });
    if (!rows.length) throw new Error('الموازنة لا تحتوي على سطور');
    const seen = {};
    let total = 0;
    const pl = rows.map(function (x, i) {
      const p = prepare_('budgets', {
        BUDGETID: id, LINENUM: i + 1, BUDGETYEAR: ph.rec.BUDGETYEAR, PERIODTYPE: ph.rec.PERIODTYPE, BUDGETTYPE: ph.rec.BUDGETTYPE,
        MAINACCOUNTID: x.MAINACCOUNTID || '', SUB2_OPERATINGUNIT: x.SUB2_OPERATINGUNIT || '', SUB1_OPERATINGUNIT: x.SUB1_OPERATINGUNIT || '', OPERATINGUNIT: x.OPERATINGUNIT || '',
        AMOUNT: toNum_(x.AMOUNT, { ar: 'سطر ' + (i + 1) + ' — المبلغ' }), NOTE: x.NOTE || ''
      }, 'add', ctx);
      const k = [low_(p.rec.MAINACCOUNTID), low_(p.rec.OPERATINGUNIT), low_(p.rec.SUB1_OPERATINGUNIT), low_(p.rec.SUB2_OPERATINGUNIT)].join('|');
      if (seen[k]) throw new Error('السطر ' + (i + 1) + ' مكرر مع السطر ' + seen[k] + ' (نفس الحساب والوحدة)');
      seen[k] = i + 1;
      total += num_(p.rec.AMOUNT);
      return p;
    });
    ph.rec.TOTALAMOUNT = round_(total, 3);
    const bt = ctx.table('budgets');
    const old = [];
    bt.rows.forEach(function (r, i) { if (!isNew && low_(r.BUDGETID) === low_(id)) old.push(i); });
    deleteRows_(bt, old);
    ctx.resetIndex('budgets');
    if (id === '@pending') {
      const no = nextSeq_('BUDGETSEQ', 'BUDGETPREFIX', 'BG-');
      ph.rec.BUDGETID = no;
      pl.forEach(function (p) { p.rec.BUDGETID = no; });
    }
    commit_('budgetheads', ph, ctx);
    pl.forEach(function (p) { commit_('budgets', p, ctx); });
    return { id: ph.rec.BUDGETID, total: ph.rec.TOTALAMOUNT, lines: pl.length };
  });
}

/** action: approve | unapprove | delete */
function budgetAction(token, id, action) {
  const s = auth_(token);
  need_(s, 'budget', action === 'delete' ? 'delete' : action === 'post' ? 'post' : 'approve');
  return withLock_(function () {
    const ctx = ctx_(s);
    const ht = ctx.table('budgetheads');
    const h = ctx.find('budgetheads', id);
    if (!h) throw new Error('الموازنة غير موجودة');
    const hi = ht.rows.indexOf(h);
    if (action === 'delete') {
      if (h.STATUS !== 'Draft') throw new Error('يمكن حذف الموازنات في حالة مسودة فقط');
      const bt = ctx.table('budgets');
      const idx = [];
      bt.rows.forEach(function (r, i) { if (low_(r.BUDGETID) === low_(id)) idx.push(i); });
      deleteRows_(bt, idx);
      deleteRows_(ht, [hi]);
      return true;
    }
    if (action === 'approve') {
      if (h.STATUS !== 'Draft') throw new Error('الموازنة ليست في حالة مسودة');
      if (!ctx.table('budgets').rows.some(function (r) { return low_(r.BUDGETID) === low_(id); })) throw new Error('الموازنة بدون سطور');
      h.STATUS = 'Approved'; h.APPROVEDBY = s.username; h.APPROVEDON = stamp_();
    } else if (action === 'unapprove') {
      if (h.STATUS !== 'Approved') throw new Error('يمكن إلغاء الاعتماد للموازنات المعتمدة فقط (غير المُرحّلة)');
      h.STATUS = 'Draft'; h.APPROVEDBY = ''; h.APPROVEDON = '';
    } else if (action === 'post') {
      if (h.STATUS !== 'Approved') throw new Error('يجب اعتماد الموازنة قبل الترحيل');
      h.STATUS = 'Posted'; h.POSTEDBY = s.username; h.POSTEDON = stamp_();
    } else throw new Error('إجراء غير معروف');
    writeRow_(ht, hi, h);
    return true;
  });
}

/* ============================== Receipts & payments ============================== */

/** قيود السند: القبض = من حـ/ الصندوق إلى حـ/ الطرف، والصرف عكسه. */
function cashLegs_(ctx, r, reverse) {
  const amt = num_(r.AMOUNT);
  const bank = { pt: 'Bank', pid: r.BANKACCOUNTID, cur: r.CURRENCYCODE, rate: num_(r.EXCHANGERATE) || 1, note: r.DESCRIPTION || '', ref: r.CHEQUENUMBER || r.VOUCHERNO };
  const party = { cur: r.CURRENCYCODE, rate: num_(r.EXCHANGERATE) || 1, s2: r.SUB2_OPERATINGUNIT || '', note: r.DESCRIPTION || '', ref: r.CHEQUENUMBER || r.VOUCHERNO };
  if (r.PARTYTYPE === 'Ledger') party.acc = r.PARTYID; else { party.pt = r.PARTYTYPE; party.pid = r.PARTYID; }
  let bankDr = r.VOUCHERTYPE === 'Receipt';
  if (reverse) bankDr = !bankDr;
  bank.curDebit = bankDr ? amt : 0; bank.curCredit = bankDr ? 0 : amt;
  party.curDebit = bankDr ? 0 : amt; party.curCredit = bankDr ? amt : 0;
  return bankDr ? [bank, party] : [party, bank];
}

function postCash_(ctx, r, reverse, date) {
  const label = r.VOUCHERTYPE === 'Receipt' ? 'سند قبض' : 'سند صرف';
  const v = autoVoucher_(ctx, { TRANSDATE: date || r.TRANSDATE, JOURNALNAME: r.VOUCHERTYPE === 'Receipt' ? 'RV' : 'PV',
    DESCRIPTION: (reverse ? 'إلغاء ' : '') + label + ' ' + r.VOUCHERNO + ' — ' + (r.PARTYNAME || r.PARTYID) + (r.DESCRIPTION ? ' — ' + r.DESCRIPTION : ''),
    REFERENCE: r.VOUCHERNO, SOURCE: 'CashVoucher', SOURCEREF: r.VOUCHERNO }, cashLegs_(ctx, r, reverse));
  writeVoucher_(ctx, v);
  return v.head.rec.VOUCHER;
}

/** حفظ سند (وترحيله مباشرة إذا post = true). */
function saveCashVoucher(token, rec, isNew, post) {
  const s = auth_(token);
  need_(s, 'cashvouchers', isNew ? 'add' : 'edit');
  if (post) need_(s, 'cashvouchers', 'post');
  return withLock_(function () {
    const ctx = ctx_(s);
    rec = Object.assign({}, rec || {});
    if (isNew) rec.VOUCHERNO = '@pending';
    else {
      const ex = ctx.find('cashvouchers', rec.VOUCHERNO);
      if (!ex) throw new Error('السند غير موجود');
      if (ex.STATUS !== 'Draft') throw new Error('لا يمكن تعديل سند مُرحّل أو ملغى');
    }
    const p = prepare_('cashvouchers', rec, isNew ? 'add' : 'edit', ctx);
    p.rec.STATUS = 'Draft';
    // نتحقق من القيد قبل أي كتابة
    if (post) prepareVoucher_(ctx, { TRANSDATE: p.rec.TRANSDATE, JOURNALNAME: p.rec.VOUCHERTYPE === 'Receipt' ? 'RV' : 'PV' }, cashLegs_(ctx, p.rec), 'add');
    if (p.rec.VOUCHERNO === '@pending') {
      p.rec.VOUCHERNO = p.rec.VOUCHERTYPE === 'Receipt' ? nextSeq_('RVSEQ', 'RVPREFIX', 'RV-') : nextSeq_('PVSEQ', 'PVPREFIX', 'PV-');
    }
    if (post) {
      p.rec.GLVOUCHER = postCash_(ctx, p.rec);
      p.rec.STATUS = 'Posted'; p.rec.POSTEDBY = s.username; p.rec.POSTEDON = stamp_();
      applyInvoicePay_(ctx, p.rec, 1);
    }
    commit_('cashvouchers', p, ctx);
    ctx.flush();
    return { no: p.rec.VOUCHERNO, status: p.rec.STATUS, gl: p.rec.GLVOUCHER || '' };
  });
}

/** action: post | cancel (قيد عكسي بتاريخ اليوم أو التاريخ المحدد) | delete (مسودة فقط) */
function cashAction(token, no, action, date) {
  const s = auth_(token);
  need_(s, 'cashvouchers', action === 'delete' ? 'delete' : 'post');
  return withLock_(function () {
    const ctx = ctx_(s);
    const t = ctx.table('cashvouchers');
    const r = ctx.find('cashvouchers', no);
    if (!r) throw new Error('السند غير موجود');
    const i = t.rows.indexOf(r);
    if (action === 'delete') {
      if (r.STATUS !== 'Draft') throw new Error('يمكن حذف السندات في حالة مسودة فقط — السند المُرحّل يُلغى');
      deleteRows_(t, [i]);
      return true;
    }
    if (action === 'post') {
      if (r.STATUS !== 'Draft') throw new Error('السند مُرحّل مسبقاً');
      checkInvoiceLink_(ctx, r);
      r.GLVOUCHER = postCash_(ctx, r);
      applyInvoicePay_(ctx, r, 1);
      r.STATUS = 'Posted'; r.POSTEDBY = s.username; r.POSTEDON = stamp_();
    } else if (action === 'cancel') {
      if (r.STATUS !== 'Posted') throw new Error('يمكن إلغاء السندات المُرحّلة فقط');
      const d = date ? toDate_(date, ctx.dateFmt, { ar: 'تاريخ الإلغاء' }) : today_();
      if (d < String(r.TRANSDATE)) throw new Error('تاريخ الإلغاء قبل تاريخ السند');
      r.REVERSALVOUCHER = postCash_(ctx, r, true, d);
      r.STATUS = 'Cancelled';
      applyInvoicePay_(ctx, r, -1);
    } else throw new Error('إجراء غير معروف');
    writeRow_(t, i, r);
    ctx.flush();
    return { gl: action === 'post' ? r.GLVOUCHER : r.REVERSALVOUCHER };
  });
}

/* ============================== Parties: customers / vendors / cash & banks ============================== */

const PARTY = {
  Customer: { entity: 'customers', id: 'CUSTOMERACCOUNT', name: 'NAME', acc: 'LEDGERACCOUNT', perm: 'customers', ar: 'العميل', sign: 1 },
  Vendor: { entity: 'vendors', id: 'VENDORACCOUNTNUMBER', name: 'VENDORORGANIZATIONNAME', acc: 'VENDORLEDGERACCOUNT', perm: 'vendors', ar: 'المورد', sign: -1 },
  Bank: { entity: 'banks', id: 'BANKACCOUNTID', name: 'NAME', acc: 'LEDGERACCOUNT', perm: 'banks', ar: 'الصندوق/البنك', sign: 1 }
};

function partyHook_(type) {
  return function (rec, ex, ctx) {
    const P = PARTY[type];
    const local = String(getSettings_().CURRENCY || '').toUpperCase();
    rec.CURRENCYCODE = String(rec.CURRENCYCODE || local).toUpperCase();
    if (rec.CURRENCYCODE !== local && !ctx.find('currencies', rec.CURRENCYCODE)) throw new Error('العملة غير معرّفة: ' + rec.CURRENCYCODE);
    rec[P.acc] = postingAccount_(ctx, rec[P.acc], 'حساب ' + P.ar);
    if (rec.BANKGROUP !== undefined) checkLk_(ctx, rec, 'BANKGROUP', 'BankGroup');
    // بعد ترحيل الرصيد الافتتاحي لا يُغيَّر إلا بعد حذف قيده
    if (ex && ex.OPENINGVOUCHER && ctx.find('glheaders', ex.OPENINGVOUCHER) && num_(ex.OPENINGBALANCE) !== num_(rec.OPENINGBALANCE)) {
      throw new Error('تم ترحيل الرصيد الافتتاحي في القيد ' + ex.OPENINGVOUCHER + ' — احذف القيد (مسودة) أو أنشئ قيد تسوية');
    }
    if (ex && ex.OPENINGVOUCHER && ctx.find('glheaders', ex.OPENINGVOUCHER) && ex.CURRENCYCODE !== rec.CURRENCYCODE) {
      throw new Error('لا يمكن تغيير العملة بعد ترحيل الرصيد الافتتاحي');
    }
    if (ex && !ctx.find('glheaders', ex.OPENINGVOUCHER || '-')) rec.OPENINGVOUCHER = '';
  };
}

function partyUsed_(ctx, type, id) {
  return ctx.table('gllines').rows.some(function (l) { return l.ACCOUNTTYPE === type && low_(l.PARTYID) === low_(id); });
}

/** بيانات تقارير الأطراف: الأطراف + سطور القيود المُرحّلة الخاصة بهم. */
function partyData(token, type) {
  const s = auth_(token);
  const P = PARTY[type];
  if (!P) throw new Error('نوع غير معروف');
  need_(s, P.perm, 'view');
  const heads = {};
  readTable_('GLJournals').rows.forEach(function (h) { if (h.VOUCHER && h.STATUS === 'Posted') heads[low_(h.VOUCHER)] = h; });
  const lines = [];
  readTable_('GLJournalLines').rows.forEach(function (l) {
    if (l.ACCOUNTTYPE !== type || !l.PARTYID) return;
    const h = heads[low_(l.VOUCHER)];
    if (!h) return;
    // [تاريخ، قيد، يومية، بيان، ملاحظة، الطرف، العملة، مدين بالعملة، دائن بالعملة، مدين محلي، دائن محلي، مرجع، سعر]
    lines.push([String(h.TRANSDATE), String(h.VOUCHER), String(h.JOURNALNAME || ''), String(h.DESCRIPTION || ''), String(l.NOTE || ''), String(l.PARTYID),
      String(l.CURRENCYCODE || ''), num_(l.AMOUNTCURDEBIT), num_(l.AMOUNTCURCREDIT), num_(l.DEBITAMOUNT), num_(l.CREDITAMOUNT),
      String(l.REFERENCE || h.REFERENCE || ''), num_(l.EXCHANGERATE) || 1]);
  });
  const parties = readTable_(SCHEMAS[P.entity].sheet).rows.filter(function (r) { return r[P.id]; }).map(function (r) {
    return { c: String(r[P.id]), n: String(r[P.name] || ''), cur: String(r.CURRENCYCODE || getSettings_().CURRENCY), acc: r[P.acc] || '',
      ob: num_(r.OPENINGBALANCE), ov: r.OPENINGVOUCHER || '', grp: r.BANKGROUP || '' };
  });
  return { type: type, parties: parties, lines: lines };
}

/** يُنشئ قيد الأرصدة الافتتاحية (مسودة) لكل الأطراف التي لم يُرحَّل رصيدها بعد. */
function createOpeningVoucher(token, type, date, offsetAccount) {
  const s = auth_(token);
  const P = PARTY[type];
  if (!P) throw new Error('نوع غير معروف');
  need_(s, P.perm, 'edit');
  need_(s, 'gl', 'add');
  return withLock_(function () {
    const ctx = ctx_(s);
    const st = getSettings_();
    const local = String(st.CURRENCY || '').toUpperCase();
    const off = postingAccount_(ctx, offsetAccount || st.OPENINGBALANCEACCOUNT, 'حساب الأرصدة الافتتاحية');
    if (!off) throw new Error('حدد حساب الأرصدة الافتتاحية (المقابل)');
    date = toDate_(date || today_(), ctx.dateFmt, { ar: 'التاريخ' });
    const t = ctx.table(P.entity);
    const todo = t.rows.filter(function (r) {
      return r[P.id] && num_(r.OPENINGBALANCE) && !(r.OPENINGVOUCHER && ctx.find('glheaders', r.OPENINGVOUCHER));
    });
    if (!todo.length) throw new Error('لا توجد أرصدة افتتاحية غير مُرحّلة');
    const legs = [];
    let localSum = 0;
    todo.forEach(function (r) {
      const cur = String(r.CURRENCYCODE || local).toUpperCase();
      const rate = cur === local ? 1 : ctx.rate(cur, date);
      if (!(rate > 0)) throw new Error('لا يوجد سعر صرف للعملة ' + cur + ' بتاريخ ' + date + ' (' + r[P.id] + ')');
      // العملاء والبنوك بطبيعة مدينة، الموردون بطبيعة دائنة
      const amt = num_(r.OPENINGBALANCE) * P.sign;
      legs.push({ pt: type, pid: r[P.id], acc: r[P.acc], cur: cur, rate: rate, curDebit: amt > 0 ? amt : 0, curCredit: amt < 0 ? -amt : 0,
        note: 'رصيد افتتاحي ' + P.ar + ' ' + r[P.id], ref: 'OPENING' });
      localSum += round_(amt * rate, 3);
    });
    localSum = round_(localSum, 3);
    if (localSum) legs.push({ acc: off, cur: local, rate: 1, curDebit: localSum < 0 ? -localSum : 0, curCredit: localSum > 0 ? localSum : 0, note: 'مقابل الأرصدة الافتتاحية', ref: 'OPENING' });
    const v = prepareVoucher_(ctx, { TRANSDATE: date, JOURNALNAME: 'GJ', DESCRIPTION: 'الأرصدة الافتتاحية — ' + P.ar, REFERENCE: 'OPENING' }, legs, 'add');
    v.head.rec.SOURCE = 'Manual';
    v.head.rec.STATUS = 'Draft';
    if (!balanced_(v)) throw new Error('قيد الأرصدة الافتتاحية غير متوازن');
    writeVoucher_(ctx, v);
    todo.forEach(function (r) { r.OPENINGVOUCHER = v.head.rec.VOUCHER; writeRow_(t, t.rows.indexOf(r), r); });
    return { voucher: v.head.rec.VOUCHER, count: todo.length };
  });
}

/* ============================== Budgets ============================== */

/** بيانات مراقبة الموازنة: سطور الموازنة للسنة + الفعلي من القيود المُرحّلة (بدون قيد الإقفال). */
function budgetData(token, year) {
  const s = auth_(token);
  need_(s, 'budget', 'view');
  year = String(year || new Date().getFullYear());
  const approved = {};
  readTable_('BudgetHeaders').rows.forEach(function (h) { if (h.BUDGETID && h.STATUS === 'Posted') approved[low_(h.BUDGETID)] = 1; });
  const budgets = readTable_('Budgets').rows.filter(function (b) { return b.LINEID && String(b.BUDGETYEAR) === year && approved[low_(b.BUDGETID)]; }).map(function (b) {
    return { acc: String(b.MAINACCOUNTID), ou: b.OPERATINGUNIT || '', s1: b.SUB1_OPERATINGUNIT || '', s2: b.SUB2_OPERATINGUNIT || '', amt: num_(b.AMOUNT), per: b.PERIODTYPE };
  });
  const heads = {};
  readTable_('GLJournals').rows.forEach(function (h) {
    if (h.VOUCHER && h.STATUS === 'Posted' && h.SOURCE !== 'YearEnd' && String(h.TRANSDATE).slice(0, 4) === year) heads[low_(h.VOUCHER)] = h;
  });
  const actual = [];
  readTable_('GLJournalLines').rows.forEach(function (l) {
    const h = heads[low_(l.VOUCHER)];
    if (!h) return;
    actual.push([String(h.TRANSDATE), String(l.MAINACCOUNTID), l.OPERATINGUNIT || '', l.SUB1_OPERATINGUNIT || '', l.SUB2_OPERATINGUNIT || '', num_(l.DEBITAMOUNT) - num_(l.CREDITAMOUNT)]);
  });
  const types = {};
  readTable_('MainAccounts').rows.forEach(function (a) { if (a.MAINACCOUNTID) types[String(a.MAINACCOUNTID)] = a.MAINACCOUNTTYPE || ''; });
  return { year: year, budgets: budgets, actual: actual, types: types };
}

/* ============================== Dashboard ============================== */

function dashboard(token) {
  const s = auth_(token);
  const out = { lookups: {} };
  const canAssets = can_(s, 'assets', 'view') || can_(s, 'reports', 'view');
  const canBooks = can_(s, 'books', 'view') || can_(s, 'reports', 'view');
  if (canAssets) {
    out.assets = { count: readTable_('FixedAssets').rows.filter(function (r) { return r.FIXEDASSETNUMBER; }).length };
  }
  if (canBooks) {
    const pb = primaryBooks_(readTable_('AssetBooks').rows);
    const t = { cost: 0, accum: 0, nbv: 0 };
    const byOU = {}, byStatus = {};
    Object.keys(pb).forEach(function (k) {
      const b = pb[k];
      const ou = b.OPERATINGUNIT || '-';
      const g = byOU[ou] = byOU[ou] || { ou: ou, cost: 0, accum: 0, nbv: 0, count: 0 };
      g.cost += num_(b.ACQUISITIONPRICE); g.accum += num_(b.ACCUMULATEDDEPRECIATION); g.nbv += nbv_(b); g.count++;
      t.cost += num_(b.ACQUISITIONPRICE); t.accum += num_(b.ACCUMULATEDDEPRECIATION); t.nbv += nbv_(b);
      byStatus[b.STATUS || '-'] = (byStatus[b.STATUS || '-'] || 0) + 1;
    });
    out.totals = { cost: round_(t.cost, 3), accum: round_(t.accum, 3), nbv: round_(t.nbv, 3) };
    out.byOU = Object.keys(byOU).map(function (k) {
      const g = byOU[k];
      return { ou: g.ou, cost: round_(g.cost, 3), accum: round_(g.accum, 3), nbv: round_(g.nbv, 3), count: g.count };
    }).sort(function (a, b) { return b.cost - a.cost; });
    out.byStatus = byStatus;
  }
  if (can_(s, 'journal', 'view')) {
    const batches = {};
    readTable_('FAJournalLines').rows.forEach(function (r) {
      if (!r.JOURNALBATCHNUMBER) return;
      const b = batches[r.JOURNALBATCHNUMBER] = batches[r.JOURNALBATCHNUMBER] || {};
      b[r.STATUS] = 1;
    });
    const c = { Draft: 0, Approved: 0, Posted: 0 };
    Object.keys(batches).forEach(function (k) {
      const b = batches[k];
      c[b.Posted ? 'Posted' : b.Approved ? 'Approved' : 'Draft']++;
    });
    out.journal = c;
  }
  return out;
}

/* ============================== Automatic depreciation ============================== */

/**
 * يحسب مقترح الإهلاك (القسط الثابت على العمر المتبقي — Straight line life remaining في D365):
 *   الإهلاك = (صافي القيمة الدفترية − قيمة الخردة) ÷ فترات الإهلاك المتبقية × عدد الشهور
 * يُحسب من الشهر التالي لآخر إهلاك حتى نهاية الفترة المحددة (عادةً 3 شهور = ربع سنة)،
 * ولا يتجاوز الفترات المتبقية، ولا ينزل بصافي القيمة الدفترية عن قيمة الخردة.
 */
function depProposal_(ctx, p) {
  const raw = toDate_(p.periodEnd || '', ctx.dateFmt, { ar: 'نهاية الفترة' });
  if (!raw) throw new Error('حدد تاريخ نهاية الفترة');
  const end = monthEnd_(raw);
  const endYM = ym_(end);
  const months = Math.max(1, Math.round(num_(p.months) || 3));
  const fGroup = low_(p.group), fOU = low_(p.ou), fBook = low_(p.book || 'General');
  const pending = {};
  ctx.table('journal').rows.forEach(function (j) {
    if (j.TRANSACTIONTYPE === 'Depreciation' && j.STATUS !== 'Posted') pending[low_(j.FIXEDASSETNUMBER) + '|' + low_(j.BOOKID)] = j.JOURNALBATCHNUMBER;
  });
  const lines = [], skipped = [];
  ctx.table('books').rows.forEach(function (b) {
    if (!b.FIXEDASSETNUMBER || low_(b.BOOKID) !== fBook) return;
    const a = ctx.asset(b.FIXEDASSETNUMBER) || {};
    if (fGroup && low_(a.FIXEDASSETGROUPID) !== fGroup) return;
    if (fOU && low_(a.OPERATINGUNIT || b.OPERATINGUNIT) !== fOU) return;
    const id = b.FIXEDASSETNUMBER;
    const skip = function (why) { skipped.push({ asset: id, name: a.NAME || '', book: b.BOOKID, reason: why }); };
    if (b.STATUS !== 'Open') return skip('حالة الدفتر ليست مفتوح');
    const price = num_(b.ACQUISITIONPRICE);
    if (price <= 0) return skip('لا يوجد سعر اقتناء');
    if (b.DEPRECIATIONPERIODSREMAINING === '' || num_(b.DEPRECIATIONPERIODSREMAINING) <= 0) return skip('لا توجد فترات إهلاك متبقية');
    if (pending[low_(id) + '|' + low_(b.BOOKID)]) return skip('يوجد إهلاك غير مُرحّل في اليومية ' + pending[low_(id) + '|' + low_(b.BOOKID)]);
    const startYM = b.LASTDEPRECIATIONDATE ? ym_(b.LASTDEPRECIATIONDATE) + 1 : ym_(b.DEPRECIATIONSTARTDATE || b.ACQUISITIONDATE);
    if (startYM === null) return skip('لا يوجد تاريخ بداية للإهلاك');
    let n = endYM - startYM + 1;
    if (n <= 0) return skip('تم إهلاكه حتى ' + b.LASTDEPRECIATIONDATE);
    const remaining = num_(b.DEPRECIATIONPERIODSREMAINING);
    n = Math.min(n, Math.ceil(remaining));
    const accum = num_(b.ACCUMULATEDDEPRECIATION);
    const nbv = round_(price - accum, 3);
    const room = round_(nbv - num_(b.SCRAPVALUE), 3);
    if (room <= 0) return skip('وصل إلى قيمة الخردة');
    // القسط الثابت على العمر المتبقي: (صافي القيمة − الخردة) ÷ الفترات المتبقية × عدد الشهور
    const monthly = room / remaining;
    const amount = n >= remaining ? room : Math.min(round_(monthly * n, 3), room);
    const g = ctx.find('groups', a.FIXEDASSETGROUPID) || {};
    lines.push({
      FIXEDASSETNUMBER: id, NAME: a.NAME || '', GROUP: a.FIXEDASSETGROUPID || '', BOOKID: b.BOOKID,
      PRICE: price, NBV: nbv, SCRAP: num_(b.SCRAPVALUE), REMAINING: remaining, MONTHLY: round_(monthly, 3),
      FROM: ymStr_(startYM), TO: end, MONTHS: n, CATCHUP: n > months, AMOUNT: -amount, ACCUM: accum,
      ACCUMAFTER: round_(accum + amount, 3), NBVAFTER: round_(nbv - amount, 3),
      LEDGER: g.ACCUMDEPACCOUNT || '', OFFSET: g.DEPEXPENSEACCOUNT || '',
      OU: a.OPERATINGUNIT || b.OPERATINGUNIT || '', SUB1: a.SUB1_OPERATINGUNIT || b.SUB1_OPERATINGUNIT || '', SUB2: a.SUB2_OPERATINGUNIT || b.SUB2_OPERATINGUNIT || ''
    });
  });
  lines.sort(function (x, y) { return String(x.FIXEDASSETNUMBER).localeCompare(String(y.FIXEDASSETNUMBER)); });
  return { periodEnd: end, months: months, lines: lines, skipped: skipped };
}

function depreciationProposal(token, p) {
  const s = auth_(token);
  need_(s, 'depreciation', 'view');
  return depProposal_(ctx_(s), p || {});
}

/** ينشئ يومية إهلاك (مسودة) من المقترح — ثم تُعتمد وتُرحّل من صفحة اليومية. */
function createDepreciationJournal(token, p) {
  const s = auth_(token);
  need_(s, 'depreciation', 'add');
  need_(s, 'journal', 'add');
  p = p || {};
  return withLock_(function () {
    const ctx = ctx_(s);
    const st = getSettings_();
    const prop = depProposal_(ctx, p);
    const only = p.only && p.only.length ? p.only.map(low_) : null;
    const lines = prop.lines.filter(function (l) { return !only || only.indexOf(low_(l.FIXEDASSETNUMBER)) >= 0; });
    if (!lines.length) throw new Error('لا توجد أصول تستحق الإهلاك في هذه الفترة');
    const jname = p.journalName || st.DEPJOURNALNAME || 'FADEP';
    const acct = p.offsetAccount || st.DEPOFFSETACCOUNT || '';
    ctx.importBatch = nextBatch_();
    let total = 0;
    lines.forEach(function (l) {
      const rec = {
        JOURNALBATCHNUMBER: ctx.importBatch, TRANSDATE: prop.periodEnd, JOURNALNAME: jname, TRANSACTIONTYPE: 'Depreciation',
        FIXEDASSETNUMBER: l.FIXEDASSETNUMBER, BOOKID: l.BOOKID, TEXT: 'إهلاك ' + l.FROM + ' → ' + prop.periodEnd.slice(0, 7) + ' (' + l.MONTHS + ' شهر)',
        DEBITAMOUNT: 0, CREDITAMOUNT: -l.AMOUNT, DEPRECIATIONMONTHS: l.MONTHS, OFFSETACCOUNT: l.OFFSET || acct,
        OFFSET_OPERATINGUNIT: l.OU, OFFSET_SUB1_OPERATINGUNIT: l.SUB1, OFFSET_SUB2_OPERATINGUNIT: l.SUB2
      };
      commit_('journal', prepare_('journal', rec, 'add', ctx), ctx);
      total += -l.AMOUNT;
    });
    return { batch: ctx.importBatch, count: lines.length, total: round_(total, 3) };
  });
}

/* ============================== Reports ============================== */

/** بيانات التقارير: الأصول مع دفاترها + حركات الإهلاك المُرحّلة. */
/** بيانات التقارير المالية: سطور القيود المُرحّلة + دليل الحسابات. */
function ledgerData(token) {
  const s = auth_(token);
  need_(s, 'reports', 'view');
  const heads = {};
  readTable_('GLJournals').rows.forEach(function (h) { if (h.VOUCHER && h.STATUS === 'Posted') heads[low_(h.VOUCHER)] = h; });
  const lines = [];
  readTable_('GLJournalLines').rows.forEach(function (l) {
    const h = heads[low_(l.VOUCHER)];
    if (!h || !l.MAINACCOUNTID) return;
    // [تاريخ، قيد، يومية، بيان، حساب، وحدة رئيسية، فرعية1، فرعية2، مدين، دائن، ملاحظة، مرجع، مصدر، سطر]
    lines.push([String(h.TRANSDATE), String(h.VOUCHER), String(h.JOURNALNAME || ''), String(h.DESCRIPTION || ''), String(l.MAINACCOUNTID),
      String(l.OPERATINGUNIT || ''), String(l.SUB1_OPERATINGUNIT || ''), String(l.SUB2_OPERATINGUNIT || ''), num_(l.DEBITAMOUNT), num_(l.CREDITAMOUNT),
      String(l.NOTE || ''), String(l.REFERENCE || h.REFERENCE || ''), String(h.SOURCE || ''), num_(l.LINENUM),
      String(l.CURRENCYCODE || ''), num_(l.AMOUNTCURDEBIT), num_(l.AMOUNTCURCREDIT)]);
  });
  const accounts = readTable_('MainAccounts').rows.filter(function (a) { return a.MAINACCOUNTID; }).map(function (a) {
    return { c: String(a.MAINACCOUNTID), n: String(a.NAME || ''), h: a.ACCOUNTLEVEL === 'Header' ? 1 : 0, p: String(a.PARENTACCOUNT || ''),
      t: a.MAINACCOUNTTYPE || '', cat: String(a.MAINACCOUNTCATEGORY || ''), cf: a.CASHFLOWCATEGORY || '' };
  });
  const fa = {};
  readTable_('AssetGroups').rows.forEach(function (g) { if (g.ASSETACCOUNT) fa[low_(g.ASSETACCOUNT)] = 1; });
  return { lines: lines, accounts: accounts, faAccounts: Object.keys(fa) };
}

function reportData(token) {
  const s = auth_(token);
  need_(s, 'reports', 'view');
  const assets = {};
  readTable_('FixedAssets').rows.forEach(function (a) { if (a.FIXEDASSETNUMBER) assets[low_(a.FIXEDASSETNUMBER)] = a; });
  const register = readTable_('AssetBooks').rows.filter(function (b) { return b.FIXEDASSETNUMBER; }).map(function (b) {
    const a = assets[low_(b.FIXEDASSETNUMBER)] || {};
    return {
      FIXEDASSETNUMBER: b.FIXEDASSETNUMBER, NAME: a.NAME || '', FIXEDASSETGROUPID: a.FIXEDASSETGROUPID || '', LOCATION: a.LOCATION || '',
      BUILDING: a.BUILDING || '', ROOMNUMBER: a.ROOMNUMBER || '', DEPARTMENT: a.DEPARTMENT || '', RESPONSIBLE: a.RESPONSIBLE || '',
      BOOKID: b.BOOKID, STATUS: b.STATUS, ACQUISITIONDATE: b.ACQUISITIONDATE, SERVICELIFE: b.SERVICELIFE,
      LASTDEPRECIATIONDATE: b.LASTDEPRECIATIONDATE, DEPRECIATIONPERIODSREMAINING: b.DEPRECIATIONPERIODSREMAINING,
      ACQUISITIONPRICE: num_(b.ACQUISITIONPRICE), ACCUMULATEDDEPRECIATION: num_(b.ACCUMULATEDDEPRECIATION), NETBOOKVALUE: nbv_(b),
      OPERATINGUNIT: a.OPERATINGUNIT || b.OPERATINGUNIT || '', SUB1_OPERATINGUNIT: a.SUB1_OPERATINGUNIT || b.SUB1_OPERATINGUNIT || '',
      SUB2_OPERATINGUNIT: a.SUB2_OPERATINGUNIT || b.SUB2_OPERATINGUNIT || '', ASSETACCOUNT: a.ASSETACCOUNT || ''
    };
  });
  const DEP = { Depreciation: 1, DepreciationAdj: 1, ExtraordinaryDepreciation: 1 };
  const depreciation = readTable_('FAJournalLines').rows.filter(function (j) { return j.STATUS === 'Posted' && DEP[j.TRANSACTIONTYPE]; })
    .map(function (j) {
      return { TRANSDATE: j.TRANSDATE, JOURNALBATCHNUMBER: j.JOURNALBATCHNUMBER, FIXEDASSETNUMBER: j.FIXEDASSETNUMBER, ASSETNAME: j.ASSETNAME,
        BOOKID: j.BOOKID, TRANSACTIONTYPE: j.TRANSACTIONTYPE, MONTHS: j.DEPRECIATIONMONTHS, AMOUNT: round_(num_(j.DEBITAMOUNT) - num_(j.CREDITAMOUNT), 3),
        OPERATINGUNIT: j.OFFSET_OPERATINGUNIT || '' };
    });
  return { register: register, depreciation: depreciation };
}

/* ============================== Users API ============================== */

function userOut_(u) {
  return {
    USERNAME: u.USERNAME, FULLNAME: u.FULLNAME, EMAIL: u.EMAIL, ROLE: u.ROLE, STATUS: u.STATUS,
    MUSTCHANGE: u.MUSTCHANGE, LASTLOGIN: u.LASTLOGIN, CREATEDON: u.CREATEDON, PERMS: effectivePerms_(u)
  };
}

function listUsers(token) {
  const s = auth_(token);
  need_(s, 'users', 'view');
  return readTable_('Users').rows.filter(function (u) { return u.USERNAME; }).map(userOut_);
}

function guardTarget_(s, target) {
  if (target && target.ROLE === 'SuperAdmin' && s.user.ROLE !== 'SuperAdmin') {
    throw new Error('لا يمكنك تعديل حساب المدير العام');
  }
}

function saveUser(token, data, isNew) {
  const s = auth_(token);
  need_(s, 'users', isNew ? 'add' : 'edit');
  data = data || {};
  const username = low_(data.USERNAME);
  if (!/^[a-z0-9._-]{3,30}$/.test(username)) throw new Error('اسم المستخدم: 3-30 حرفاً إنجليزياً أو أرقاماً أو . _ -');
  const role = ROLES[data.ROLE] ? data.ROLE : 'Viewer';
  if (role === 'SuperAdmin' && s.user.ROLE !== 'SuperAdmin') throw new Error('فقط المدير العام يمكنه منح دور المدير العام');
  const email = String(data.EMAIL || '').trim();
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('بريد إلكتروني غير صحيح');
  const fullName = String(data.FULLNAME || '').trim();
  if (!fullName) throw new Error('الاسم الكامل مطلوب');

  // لا يمكن منح صلاحية لا يملكها من يمنحها
  const perms = {};
  PERM_PAGES.forEach(function (pg) {
    perms[pg] = {};
    PERM_ACTIONS.forEach(function (a) {
      const want = !!(data.PERMS && data.PERMS[pg] && data.PERMS[pg][a]);
      perms[pg][a] = want && (s.user.ROLE === 'SuperAdmin' || can_(s, pg, a));
    });
  });

  return withLock_(function () {
    const f = findUser_(username);
    const now = stamp_();
    if (isNew) {
      if (f) throw new Error('اسم المستخدم موجود مسبقاً');
      checkPassword_(data.PASSWORD);
      const salt = Utilities.getUuid();
      const t = readTable_('Users');
      appendRow_(t, {
        USERNAME: username, FULLNAME: fullName, EMAIL: email, PASSWORDHASH: hash_(data.PASSWORD, salt), SALT: salt,
        ROLE: role, PERMISSIONS: JSON.stringify(perms), STATUS: data.STATUS === 'Inactive' ? 'Inactive' : 'Active',
        MUSTCHANGE: 'TRUE', LASTLOGIN: '', CREATEDBY: s.username, CREATEDON: now, MODIFIEDBY: '', MODIFIEDON: ''
      });
    } else {
      if (!f) throw new Error('المستخدم غير موجود');
      guardTarget_(s, f.u);
      if (low_(f.u.USERNAME) === low_(s.username) && role !== f.u.ROLE) throw new Error('لا يمكنك تغيير دورك بنفسك');
      if (f.u.ROLE === 'SuperAdmin' && role !== 'SuperAdmin') ensureAnotherSuperAdmin_(f.t, username);
      f.u.FULLNAME = fullName; f.u.EMAIL = email; f.u.ROLE = role; f.u.PERMISSIONS = JSON.stringify(perms);
      f.u.MODIFIEDBY = s.username; f.u.MODIFIEDON = now;
      writeRow_(f.t, f.idx, f.u);
      clearUserCache_(username);
    }
    return true;
  });
}

function ensureAnotherSuperAdmin_(t, username) {
  const others = t.rows.filter(function (u) { return u.ROLE === 'SuperAdmin' && u.STATUS === 'Active' && low_(u.USERNAME) !== low_(username); });
  if (!others.length) throw new Error('يجب أن يبقى مدير عام نشط واحد على الأقل');
}

function setUserStatus(token, username, status) {
  const s = auth_(token);
  need_(s, 'users', 'edit');
  return withLock_(function () {
    const f = findUser_(username);
    if (!f) throw new Error('المستخدم غير موجود');
    guardTarget_(s, f.u);
    if (low_(f.u.USERNAME) === low_(s.username)) throw new Error('لا يمكنك إيقاف حسابك');
    if (status !== 'Active' && f.u.ROLE === 'SuperAdmin') ensureAnotherSuperAdmin_(f.t, username);
    f.u.STATUS = status === 'Active' ? 'Active' : 'Inactive';
    f.u.MODIFIEDBY = s.username; f.u.MODIFIEDON = stamp_();
    writeRow_(f.t, f.idx, f.u);
    clearUserCache_(username);
    if (f.u.STATUS !== 'Active') endUserSessions_(username);
    return true;
  });
}

function resetUserPassword(token, username, newPw) {
  const s = auth_(token);
  need_(s, 'users', 'edit');
  checkPassword_(newPw);
  return withLock_(function () {
    const f = findUser_(username);
    if (!f) throw new Error('المستخدم غير موجود');
    guardTarget_(s, f.u);
    const salt = Utilities.getUuid();
    f.u.SALT = salt; f.u.PASSWORDHASH = hash_(newPw, salt); f.u.MUSTCHANGE = 'TRUE';
    f.u.MODIFIEDBY = s.username; f.u.MODIFIEDON = stamp_();
    writeRow_(f.t, f.idx, f.u);
    clearUserCache_(username);
    endUserSessions_(username);
    return true;
  });
}

function deleteUser(token, username) {
  const s = auth_(token);
  need_(s, 'users', 'delete');
  return withLock_(function () {
    const f = findUser_(username);
    if (!f) throw new Error('المستخدم غير موجود');
    guardTarget_(s, f.u);
    if (low_(f.u.USERNAME) === low_(s.username)) throw new Error('لا يمكنك حذف حسابك');
    if (f.u.ROLE === 'SuperAdmin') ensureAnotherSuperAdmin_(f.t, username);
    endUserSessions_(username);
    deleteRows_(f.t, [f.idx]);
    clearUserCache_(username);
    return true;
  });
}

/* ============================== Settings API ============================== */

function saveSettings(token, data) {
  const s = auth_(token);
  need_(s, 'settings', 'edit');
  data = data || {};
  const out = {};
  if ('ORGNAME' in data) out.ORGNAME = String(data.ORGNAME).trim().slice(0, 120);
  if ('ORGEMAIL' in data) {
    const e = String(data.ORGEMAIL).trim();
    if (e && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)) throw new Error('بريد إلكتروني غير صحيح');
    out.ORGEMAIL = e;
  }
  if ('CURRENCY' in data) {
    const c = String(data.CURRENCY).trim().toUpperCase();
    if (!/^[A-Z]{3}$/.test(c)) throw new Error('رمز العملة يجب أن يكون 3 أحرف (مثل DJF أو USD)');
    out.CURRENCY = c;
  }
  if ('SESSIONTIMEOUT' in data) {
    const n = Math.round(Number(data.SESSIONTIMEOUT));
    if (!(n >= 5 && n <= 1440)) throw new Error('مدة الجلسة بين 5 و 1440 دقيقة');
    out.SESSIONTIMEOUT = n;
  }
  if ('MAXSESSIONS' in data) {
    const n = Math.round(Number(data.MAXSESSIONS));
    if (!(n >= 1 && n <= 20)) throw new Error('عدد الجلسات المتزامنة بين 1 و 20');
    out.MAXSESSIONS = n;
  }
  if ('DATEFORMAT' in data) out.DATEFORMAT = data.DATEFORMAT === 'DMY' ? 'DMY' : 'MDY';
  if ('RETAINEDEARNINGSACCOUNT' in data) out.RETAINEDEARNINGSACCOUNT = String(data.RETAINEDEARNINGSACCOUNT).trim().slice(0, 40);
  if ('OPENINGBALANCEACCOUNT' in data) out.OPENINGBALANCEACCOUNT = String(data.OPENINGBALANCEACCOUNT).trim().slice(0, 40);
  ['SALESACCOUNT', 'PURCHASEACCOUNT', 'SALESTAXACCOUNT', 'PURCHTAXACCOUNT', 'SALESRETURNACCOUNT', 'PURCHRETURNACCOUNT'].forEach(function (k) { if (k in data) out[k] = String(data[k]).trim().slice(0, 40); });
  if ('DEFAULTTAXPERCENT' in data) {
    const tx = Number(data.DEFAULTTAXPERCENT);
    if (!(tx >= 0 && tx <= 100)) throw new Error('نسبة الضريبة الافتراضية بين 0 و100');
    out.DEFAULTTAXPERCENT = String(tx);
  }
  if ('DEPJOURNALNAME' in data) out.DEPJOURNALNAME = String(data.DEPJOURNALNAME).trim().slice(0, 20);
  if ('DEPOFFSETACCOUNT' in data) out.DEPOFFSETACCOUNT = String(data.DEPOFFSETACCOUNT).trim().slice(0, 40);
  if ('DEPMONTHS' in data) {
    const n = Math.round(Number(data.DEPMONTHS));
    if (!(n >= 1 && n <= 12)) throw new Error('عدد شهور دورة الإهلاك بين 1 و 12');
    out.DEPMONTHS = n;
  }
  if ('JOURNALPREFIX' in data) {
    const p = String(data.JOURNALPREFIX).trim();
    if (!/^[A-Za-z0-9_-]{1,10}$/.test(p)) throw new Error('بادئة اليومية: حتى 10 أحرف إنجليزية/أرقام');
    out.JOURNALPREFIX = p;
  }
  withLock_(function () { setSettings_(out); });
  return publicSettings_();
}

function saveLogo(token, dataUrl) {
  const s = auth_(token);
  need_(s, 'settings', 'edit');
  dataUrl = String(dataUrl || '');
  if (dataUrl && !/^data:image\/(png|jpeg|webp|gif);base64,[A-Za-z0-9+\/=]+$/.test(dataUrl)) throw new Error('صيغة الصورة غير مدعومة');
  if (dataUrl.length > 48000) throw new Error('حجم الشعار كبير، استخدم صورة أصغر');
  withLock_(function () { setSettings_({ LOGO: dataUrl }); });
  return true;
}

function listSessions(token) {
  const s = auth_(token);
  need_(s, 'settings', 'view');
  const timeoutMs = Number(getSettings_().SESSIONTIMEOUT || 60) * 60000;
  const now = Date.now();
  const users = {};
  readTable_('Users').rows.forEach(function (u) { users[u.USERNAME] = u.FULLNAME; });
  return readTable_('Sessions').rows.filter(function (r) {
    return isTrue_(r.ACTIVE) && now - (Date.parse(r.LASTSEEN) || 0) <= timeoutMs;
  }).map(function (r) {
    return { id: r.SESSIONID, username: r.USERNAME, fullName: users[r.USERNAME] || '', created: r.CREATEDON,
      lastSeen: r.LASTSEEN, current: r.SESSIONID === s.sid };
  });
}

function endSessionById(token, id) {
  const s = auth_(token);
  need_(s, 'settings', 'edit');
  if (id === s.sid) throw new Error('لا يمكنك إنهاء جلستك الحالية من هنا، استخدم تسجيل الخروج');
  withLock_(function () { endSession_(id); });
  return true;
}

/** تنظيف الجلسات المنتهية — يمكن ربطها بمشغّل زمني (Trigger) يومي. */
function cleanupSessions() {
  withLock_(function () {
    const t = readTable_('Sessions');
    const limit = Date.now() - 7 * 86400000;
    const idxs = [];
    t.rows.forEach(function (r, i) { if (!isTrue_(r.ACTIVE) || (Date.parse(r.LASTSEEN) || 0) < limit) idxs.push(i); });
    deleteRows_(t, idxs);
  });
}
