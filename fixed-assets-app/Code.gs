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

const APP_VERSION = '5.0.0';
const SETUP_VERSION = '5';
const SS_PROP = 'FA_SPREADSHEET_ID';
const HASH_ROUNDS = 300;
const AUTH_ERR = 'AUTH: ';
const AUDIT = ['CREATEDBY', 'CREATEDON', 'MODIFIEDBY', 'MODIFIEDON'];
const PERM_PAGES = ['assets', 'books', 'journal', 'transfers', 'gl', 'depreciation', 'reports', 'masterdata', 'vendors', 'users', 'settings'];
const PERM_ACTIONS = ['view', 'add', 'edit', 'delete', 'approve', 'post'];
const USER_COLS = ['USERNAME', 'FULLNAME', 'EMAIL', 'PASSWORDHASH', 'SALT', 'ROLE', 'PERMISSIONS', 'STATUS',
  'MUSTCHANGE', 'LASTLOGIN', 'CREATEDBY', 'CREATEDON', 'MODIFIEDBY', 'MODIFIEDON'];
const SESSION_COLS = ['SESSIONID', 'USERNAME', 'CREATEDON', 'LASTSEEN', 'ACTIVE', 'CLIENT'];
const SETTING_DEFAULTS = {
  ORGNAME: 'إدارة الأصول الثابتة', ORGEMAIL: '', CURRENCY: 'DJF', SESSIONTIMEOUT: '60',
  MAXSESSIONS: '3', DATEFORMAT: 'MDY', LOGO: '', JOURNALPREFIX: 'FAJ-', JOURNALSEQ: '0', TRANSFERPREFIX: 'TRF-', TRANSFERSEQ: '0', GLPREFIX: 'JV-', GLSEQ: '0', TRFJOURNALNAME: 'FAJ',
  DEPJOURNALNAME: 'FADEP', DEPOFFSETACCOUNT: '', DEPMONTHS: '3'
};
const PUBLIC_SETTINGS = ['ORGNAME', 'ORGEMAIL', 'CURRENCY', 'SESSIONTIMEOUT', 'MAXSESSIONS', 'DATEFORMAT', 'LOGO', 'JOURNALPREFIX',
  'DEPJOURNALNAME', 'DEPOFFSETACCOUNT', 'DEPMONTHS', 'TRANSFERPREFIX', 'GLPREFIX', 'TRFJOURNALNAME'];

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
const GL_SOURCES = [o('Manual', 'يدوي', 'Manual'), o('FAJournal', 'يومية الأصول', 'Fixed asset journal'), o('Transfer', 'نقل أصول', 'Asset transfer')];

const ACCOUNT_TYPES = [o('ProfitAndLoss', 'أرباح وخسائر', 'Profit and loss'), o('Revenue', 'إيرادات', 'Revenue'),
  o('Expense', 'مصروفات', 'Expense'), o('BalanceSheet', 'الميزانية العمومية', 'Balance sheet'), o('Asset', 'أصول', 'Asset'),
  o('Liability', 'خصوم', 'Liability'), o('Equity', 'حقوق ملكية', 'Equity'), o('Total', 'إجمالي', 'Total'),
  o('Reporting', 'تقارير', 'Reporting')];
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
  { v: 'JournalName', sheet: 'JournalNames', ar: 'أسماء اليوميات', en: 'Journal names' }
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
        aliases: ['Default dimension', 'DEFAULTDIMENSION', 'Financial dimensions'] })
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
      F('MAINACCOUNTID', 'رقم الحساب', 'Main account', 'lookup', { lk: '@accounts', req: 1, sec: 'l', grp: 'l1', list: 1, aliases: ['Account', 'MAINACCOUNT'] }),
      F('ACCOUNTNAME', 'اسم الحساب', 'Account name', 'text', { calc: 1, sec: 'l', grp: 'l1', list: 1 })
    ].concat(unitFields_('', 'l', 'l1')).concat([
      F('DEBITAMOUNT', 'مدين', 'Debit', 'number', { dec: 3, sec: 'l', grp: 'l2', list: 1 }),
      F('CREDITAMOUNT', 'دائن', 'Credit', 'number', { dec: 3, sec: 'l', grp: 'l2', list: 1 }),
      F('NOTE', 'ملاحظة', 'Note', 'text', { sec: 'l', grp: 'l2', list: 1, aliases: ['Text', 'TEXT'] }),
      F('REFERENCE', 'المرجع', 'Reference', 'text', { sec: 'l', grp: 'l2', list: 1 }),
      F('LEDGERDIMENSIONDISPLAYVALUE', 'الحساب والأبعاد', 'Ledger account display value', 'text', { calc: 1, sec: 'l', grp: 'l1', list: 1,
        aliases: ['ACCOUNTDISPLAYVALUE'] })
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
        aliases: ['Summary account', 'Ledger account', 'Payable account'] })
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
    .setTitle(getSettings_().ORGNAME || 'Fixed Assets')
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
  const data = ['assets', 'books', 'journal', 'vendors', 'masterdata'];
  if (role === 'SuperAdmin' || role === 'Admin') set(PERM_PAGES, PERM_ACTIONS);
  else if (role === 'Accountant') { set(data.concat(['transfers', 'gl']), ['view', 'add', 'edit']); set(['depreciation'], ['view', 'add']); set(['reports'], ['view']); }
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
  const out = { assets: [], vendors: [], books: [], accounts: [], groups: [] };
  out.accounts = readTable_('MainAccounts').rows.filter(function (r) { return r.MAINACCOUNTID; })
    .map(function (r) { return { c: String(r.MAINACCOUNTID), n: String(r.NAME || ''), p: String(r.PARENTACCOUNT || ''), h: r.ACCOUNTLEVEL === 'Header' ? 1 : 0 }; });
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
  if (can_(s, 'vendors', 'view') || can_(s, 'journal', 'view')) {
    out.vendors = readTable_('Vendors').rows.filter(function (r) { return r.VENDORACCOUNTNUMBER; })
      .map(function (r) { return { c: String(r.VENDORACCOUNTNUMBER), n: String(r.VENDORORGANIZATIONNAME || ''), acc: r.VENDORLEDGERACCOUNT || '' }; });
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

  vendors: function (rec, ex, ctx) {
    rec.VENDORLEDGERACCOUNT = postingAccount_(ctx, rec.VENDORLEDGERACCOUNT, 'حساب المورد');
  },

  journal: function (rec, ex, ctx) {
    if (ex && ex.STATUS && ex.STATUS !== 'Draft') throw new Error('لا يمكن تعديل سطر معتمد أو مُرحّل');
    if (ex) rec.JOURNALBATCHNUMBER = ex.JOURNALBATCHNUMBER;
    if (!rec.JOURNALBATCHNUMBER) rec.JOURNALBATCHNUMBER = ctx.batch();
    if (!ex) {
      const st = ctx.batchStatus(rec.JOURNALBATCHNUMBER);
      if (st && st !== 'Draft') throw new Error('لا يمكن إضافة أسطر إلى يومية معتمدة أو مُرحّلة: ' + rec.JOURNALBATCHNUMBER);
    }
    checkLk_(ctx, rec, 'JOURNALNAME', 'JournalName');
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
  if (!rec.STATUS) rec.STATUS = 'Draft';
};

HOOKS.gllines = function (rec, ex, ctx) {
  rec.MAINACCOUNTID = postingAccount_(ctx, rec.MAINACCOUNTID, 'سطر ' + (rec.LINENUM || '') + ' — رقم الحساب');
  rec.ACCOUNTNAME = (ctx.find('accounts', rec.MAINACCOUNTID) || {}).NAME || '';
  deriveChain_(ctx, rec, UNIT_CHAIN);
  const d = num_(rec.DEBITAMOUNT), c = num_(rec.CREDITAMOUNT);
  if (d < 0 || c < 0) throw new Error('سطر ' + rec.LINENUM + ': المبالغ لا يمكن أن تكون سالبة');
  if (d > 0 && c > 0) throw new Error('سطر ' + rec.LINENUM + ': أدخل مديناً أو دائناً فقط');
  if (!d && !c) throw new Error('سطر ' + rec.LINENUM + ': أدخل مبلغاً');
  rec.DEBITAMOUNT = d; rec.CREDITAMOUNT = c;
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
      VOUCHER: ph.rec.VOUCHER, LINENUM: i + 1, MAINACCOUNTID: x.acc, OPERATINGUNIT: x.ou || '', SUB1_OPERATINGUNIT: x.s1 || '',
      SUB2_OPERATINGUNIT: x.s2 || '', DEBITAMOUNT: x.debit || 0, CREDITAMOUNT: x.credit || 0, NOTE: x.note || '', REFERENCE: x.ref || ''
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
      return x && (x.MAINACCOUNTID || num_(x.DEBITAMOUNT) || num_(x.CREDITAMOUNT));
    }).map(function (x) {
      return { acc: x.MAINACCOUNTID, s2: x.SUB2_OPERATINGUNIT, s1: x.SUB1_OPERATINGUNIT, ou: x.OPERATINGUNIT,
        debit: toNum_(x.DEBITAMOUNT, { ar: 'مدين' }) || 0, credit: toNum_(x.CREDITAMOUNT, { ar: 'دائن' }) || 0, note: x.NOTE, ref: x.REFERENCE };
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
      h.STATUS = 'Posted'; h.POSTEDBY = s.username; h.POSTEDON = now;
    } else {
      throw new Error('إجراء غير معروف');
    }
    writeRow_(ht, hi, h);
    return true;
  });
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
