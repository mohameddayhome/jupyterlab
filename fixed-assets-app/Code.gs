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

const APP_VERSION = '2.0.0';
const SETUP_VERSION = '2';
const SS_PROP = 'FA_SPREADSHEET_ID';
const HASH_ROUNDS = 300;
const AUTH_ERR = 'AUTH: ';
const AUDIT = ['CREATEDBY', 'CREATEDON', 'MODIFIEDBY', 'MODIFIEDON'];
const PERM_PAGES = ['assets', 'books', 'journal', 'depreciation', 'reports', 'masterdata', 'vendors', 'users', 'settings'];
const PERM_ACTIONS = ['view', 'add', 'edit', 'delete', 'approve', 'post'];
const USER_COLS = ['USERNAME', 'FULLNAME', 'EMAIL', 'PASSWORDHASH', 'SALT', 'ROLE', 'PERMISSIONS', 'STATUS',
  'MUSTCHANGE', 'LASTLOGIN', 'CREATEDBY', 'CREATEDON', 'MODIFIEDBY', 'MODIFIEDON'];
const SESSION_COLS = ['SESSIONID', 'USERNAME', 'CREATEDON', 'LASTSEEN', 'ACTIVE', 'CLIENT'];
const SETTING_DEFAULTS = {
  ORGNAME: 'إدارة الأصول الثابتة', ORGEMAIL: '', CURRENCY: 'DJF', SESSIONTIMEOUT: '60',
  MAXSESSIONS: '3', DATEFORMAT: 'MDY', LOGO: '', JOURNALPREFIX: 'FAJ-', JOURNALSEQ: '0',
  DEPJOURNALNAME: 'FADEP', DEPOFFSETACCOUNT: '', DEPMONTHS: '3'
};
const PUBLIC_SETTINGS = ['ORGNAME', 'ORGEMAIL', 'CURRENCY', 'SESSIONTIMEOUT', 'MAXSESSIONS', 'DATEFORMAT', 'LOGO', 'JOURNALPREFIX',
  'DEPJOURNALNAME', 'DEPOFFSETACCOUNT', 'DEPMONTHS'];

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
const OFFSET_TYPES = [o('Ledger', 'دفتر الأستاذ', 'Ledger'), o('Vendor', 'مورد', 'Vendor')];

const LOOKUP_TYPES = [
  o('AssetGroup', 'مجموعات الأصول', 'Fixed asset groups'), o('Location', 'المواقع', 'Locations'),
  o('Department', 'الإدارات / الأقسام', 'Departments'), o('OperatingUnit', 'الوحدة التشغيلية', 'A_OperatingUnit'),
  o('Sub1_OperatingUnit', 'الوحدة الفرعية 1', 'A_Sub1_OperatingUnit'),
  o('Sub2_OperatingUnit', 'الوحدة الفرعية 2', 'A_Sub2_OperatingUnit'),
  o('JournalName', 'أسماء اليوميات', 'Journal names'), o('Building', 'المباني', 'Buildings'),
  o('Room', 'الغرف', 'Rooms'), o('Employee', 'الموظفون المسؤولون', 'Responsible employees')
];

/** أنواع القوائم التي لها أب (الكود الأب يُختار من النوع المذكور). */
const LOOKUP_PARENTS = { Sub1_OperatingUnit: 'OperatingUnit', Sub2_OperatingUnit: 'Sub1_OperatingUnit', Room: 'Building' };

const ACCOUNT_TYPES = [o('ProfitAndLoss', 'أرباح وخسائر', 'Profit and loss'), o('Revenue', 'إيرادات', 'Revenue'),
  o('Expense', 'مصروفات', 'Expense'), o('BalanceSheet', 'الميزانية العمومية', 'Balance sheet'), o('Asset', 'أصول', 'Asset'),
  o('Liability', 'خصوم', 'Liability'), o('Equity', 'حقوق ملكية', 'Equity'), o('Total', 'إجمالي', 'Total'),
  o('Reporting', 'تقارير', 'Reporting')];

const SCHEMAS = {
  assets: {
    sheet: 'FixedAssets', perm: 'assets', key: ['FIXEDASSETNUMBER'], title: { ar: 'الأصول الثابتة', en: 'Fixed assets' },
    sections: [
      { id: 'general', ar: 'عام', en: 'General', groups: [
        { id: 'id', ar: 'التعريف', en: 'Identification' }, { id: 'desc', ar: 'الوصف', en: 'Description' },
        { id: 'type', ar: 'النوع', en: 'Type' }, { id: 'qty', ar: 'الكمية', en: 'Quantity' }] },
      { id: 'tech', ar: 'معلومات فنية', en: 'Technical information', groups: [
        { id: 'model', ar: 'الموديل', en: 'Model' }] },
      { id: 'loc', ar: 'الموقع والمسؤولية', en: 'Location & responsibility', groups: [
        { id: 'loc', ar: 'الموقع', en: 'Location' }, { id: 'emp', ar: 'المسؤول', en: 'Responsible' }] }
    ],
    fields: [
      F('FIXEDASSETNUMBER', 'رقم الأصل', 'Number', 'text', { req: 1, key: 1, sec: 'general', grp: 'id', list: 1, sum: 1,
        aliases: ['Fixed asset number', 'Asset ID', 'ASSETID', 'Fixed asset'] }),
      F('FIXEDASSETGROUPID', 'مجموعة الأصل', 'Fixed asset group', 'lookup', { lk: 'AssetGroup', req: 1, sec: 'general', grp: 'id', list: 1,
        aliases: ['Fixed asset group', 'Asset group', 'ASSETGROUP', 'GROUPID'] }),
      F('NAME', 'اسم الأصل', 'Name', 'text', { req: 1, sec: 'general', grp: 'desc', list: 1, sum: 1, aliases: ['Asset name', 'Fixed asset name'] }),
      F('ASSETTYPE', 'النوع', 'Type', 'select', { opts: ASSET_TYPES, dflt: 'Tangible', sec: 'general', grp: 'type', list: 1, sum: 1,
        aliases: ['Asset type', 'TYPE'] }),
      F('QUANTITY', 'الكمية', 'Quantity', 'number', { dec: 2, dflt: 1, sec: 'general', grp: 'qty', list: 1 }),
      F('MODEL', 'الموديل', 'Model', 'text', { sec: 'tech', grp: 'model', list: 1, sum: 1 }),
      F('GUARANTEEDATE', 'تاريخ الضمان', 'Guarantee date', 'date', { sec: 'tech', grp: 'model', sum: 1,
        aliases: ['Warranty date', 'WARRANTYDATE'] }),
      F('LOCATION', 'الموقع', 'Location', 'lookup', { lk: 'Location', sec: 'loc', grp: 'loc', list: 1, sum: 1,
        aliases: ['Location ID', 'LOCATIONID', 'Location memo', 'LOCATIONMEMO'] }),
      F('DEPARTMENT', 'الإدارة / القسم', 'Department', 'lookup', { lk: 'Department', sec: 'loc', grp: 'loc', list: 1 }),
      F('BUILDING', 'المبنى', 'Building', 'lookup', { lk: 'Building', sec: 'loc', grp: 'loc', list: 1, aliases: ['Building ID'] }),
      F('ROOMNUMBER', 'رقم الغرفة', 'Room number', 'lookup', { lk: 'Room', parent: 'BUILDING', sec: 'loc', grp: 'loc', aliases: ['Room'] }),
      F('RESPONSIBLE', 'الموظف المسؤول', 'Employee', 'lookup', { lk: 'Employee', sec: 'loc', grp: 'emp', list: 1, sum: 1,
        aliases: ['Worker responsible', 'WORKERRESPONSIBLE', 'Responsible', 'Employee', 'WORKERRESPONSIBLEPERSONNELNUMBER'] })
    ]
  },

  books: {
    sheet: 'AssetBooks', perm: 'books', key: ['FIXEDASSETNUMBER', 'BOOKID'], title: { ar: 'دفاتر الأصول', en: 'Books' },
    sections: [
      { id: 'general', ar: 'عام', en: 'General', groups: [
        { id: 'book', ar: 'الدفتر', en: 'Book' }, { id: 'acq', ar: 'الاقتناء', en: 'Acquisition' },
        { id: 'bal', ar: 'الرصيد', en: 'Balance' }] },
      { id: 'dep', ar: 'الإهلاك', en: 'Depreciation', groups: [
        { id: 'life', ar: 'العمر الإنتاجي', en: 'Service life' }, { id: 'run', ar: 'الإهلاك', en: 'Depreciation' }] },
      { id: 'dims', ar: 'الأبعاد المالية الافتراضية', en: 'Default financial dimensions', groups: [
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
        aliases: ['Last depreciation date', 'LASTDEPRECIATIONRUNDATE'] }),
      F('OPERATINGUNIT', 'الوحدة التشغيلية', 'A_A_OperatingUnit', 'lookup', { lk: 'OperatingUnit', sec: 'dims', grp: 'dims', list: 1,
        aliases: ['A_OperatingUnit', 'A_A_OperatingUnit', 'OperatingUnit'] }),
      F('SUB1_OPERATINGUNIT', 'الوحدة الفرعية 1', 'A_Sub1_OperatingUnit', 'lookup', { lk: 'Sub1_OperatingUnit', parent: 'OPERATINGUNIT', sec: 'dims', grp: 'dims',
        aliases: ['A_Sub1_OperatingUnit', 'Sub1_OperatingUnit'] }),
      F('SUB2_OPERATINGUNIT', 'الوحدة الفرعية 2', 'A_Sub2_OperatingUnit', 'lookup', { lk: 'Sub2_OperatingUnit', parent: 'SUB1_OPERATINGUNIT', sec: 'dims', grp: 'dims',
        aliases: ['A_Sub2_OperatingUnit', 'Sub2_OperatingUnit'] }),
      F('DEFAULTDIMENSIONDISPLAYVALUE', 'قيمة الأبعاد الافتراضية', 'Default dimension display value', 'text', { calc: 1, sec: 'dims', grp: 'dims',
        aliases: ['Default dimension', 'DEFAULTDIMENSION', 'Financial dimensions'] })
    ]
  },

  journal: {
    sheet: 'FAJournalLines', perm: 'journal', key: ['LINEID'], autoId: 'LINEID', title: { ar: 'يومية الأصول الثابتة', en: 'Fixed asset journal' },
    sections: [
      { id: 'line', ar: 'السطر', en: 'Line', groups: [
        { id: 'batch', ar: 'اليومية', en: 'Journal' }, { id: 'asset', ar: 'الأصل', en: 'Fixed asset' },
        { id: 'amt', ar: 'المبلغ', en: 'Amount' }] },
      { id: 'offset', ar: 'الحساب المقابل', en: 'Offset account', groups: [
        { id: 'offset', ar: 'الحساب المقابل', en: 'Offset account' }] },
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
      F('TEXT', 'الوصف', 'Description', 'text', { sec: 'line', grp: 'amt', list: 1, aliases: ['Description', 'TRANSTEXT', 'Transaction text'] }),
      F('DEBITAMOUNT', 'مدين', 'Debit', 'number', { dec: 3, sec: 'line', grp: 'amt', list: 1, aliases: ['AMOUNTCURDEBIT', 'Debit amount'] }),
      F('CREDITAMOUNT', 'دائن', 'Credit', 'number', { dec: 3, sec: 'line', grp: 'amt', list: 1, aliases: ['AMOUNTCURCREDIT', 'Credit amount'] }),
      F('DEPRECIATIONMONTHS', 'عدد شهور الإهلاك', 'Depreciation months', 'number', { dec: 0, sec: 'line', grp: 'amt',
        aliases: ['Depreciation periods', 'PERIODS'] }),
      F('OFFSETACCOUNTTYPE', 'نوع الحساب المقابل', 'Offset account type', 'select', { opts: OFFSET_TYPES, dflt: 'Ledger', sec: 'offset', grp: 'offset' }),
      F('OFFSETACCOUNT', 'رقم الحساب', 'Account number', 'lookup', { lkBy: 'OFFSETACCOUNTTYPE', lkMap: { Ledger: '@accounts', Vendor: '@vendors' },
        sec: 'offset', grp: 'offset', aliases: ['Offset main account', 'OFFSETMAINACCOUNT', 'MAINACCOUNT'] }),
      F('OFFSET_OPERATINGUNIT', 'الوحدة التشغيلية', 'A_A_OperatingUnit', 'lookup', { lk: 'OperatingUnit', sec: 'offset', grp: 'offset' }),
      F('OFFSET_SUB1_OPERATINGUNIT', 'الوحدة الفرعية 1', 'A_Sub1_OperatingUnit', 'lookup', { lk: 'Sub1_OperatingUnit', parent: 'OFFSET_OPERATINGUNIT', sec: 'offset', grp: 'offset' }),
      F('OFFSET_SUB2_OPERATINGUNIT', 'الوحدة الفرعية 2', 'A_Sub2_OperatingUnit', 'lookup', { lk: 'Sub2_OperatingUnit', parent: 'OFFSET_SUB1_OPERATINGUNIT', sec: 'offset', grp: 'offset' }),
      F('OFFSETACCOUNTDISPLAYVALUE', 'الحساب المقابل', 'Offset account', 'text', { calc: 1, sec: 'offset', grp: 'offset', list: 1,
        aliases: ['Offset account display value'] }),
      F('STATUS', 'الحالة', 'Status', 'select', { opts: JOURNAL_STATUS, sys: 1, sec: 'wf', grp: 'wf', list: 1 }),
      F('APPROVEDBY', 'اعتمد بواسطة', 'Approved by', 'text', { sys: 1, sec: 'wf', grp: 'wf' }),
      F('APPROVEDON', 'تاريخ الاعتماد', 'Approved on', 'text', { sys: 1, sec: 'wf', grp: 'wf' }),
      F('POSTEDBY', 'رحّل بواسطة', 'Posted by', 'text', { sys: 1, sec: 'wf', grp: 'wf' }),
      F('POSTEDON', 'تاريخ الترحيل', 'Posted on', 'text', { sys: 1, sec: 'wf', grp: 'wf' })
    ]
  },

  vendors: {
    sheet: 'Vendors', perm: 'vendors', key: ['VENDORACCOUNTNUMBER'], title: { ar: 'الموردون', en: 'Vendors' },
    sections: [{ id: 'general', ar: 'عام', en: 'General', groups: [
      { id: 'id', ar: 'التعريف', en: 'Identification' }, { id: 'contact', ar: 'الاتصال', en: 'Contact' }] }],
    fields: [
      F('VENDORACCOUNTNUMBER', 'رقم المورد', 'Vendor account', 'text', { req: 1, key: 1, sec: 'general', grp: 'id', list: 1,
        aliases: ['Vendor account number', 'Vendor', 'ACCOUNTNUM'] }),
      F('VENDORORGANIZATIONNAME', 'اسم المورد', 'Name', 'text', { req: 1, sec: 'general', grp: 'id', list: 1,
        aliases: ['Vendor name', 'NAME', 'ORGANIZATIONNAME'] }),
      F('ADDRESS', 'العنوان', 'Address', 'textarea', { sec: 'general', grp: 'contact', list: 1,
        aliases: ['FULLPRIMARYADDRESS', 'Primary address', 'ADDRESSDESCRIPTION', 'ADDRESSSTREET'] }),
      F('PRIMARYEMAILADDRESS', 'البريد الإلكتروني', 'Email', 'email', { sec: 'general', grp: 'contact', list: 1,
        aliases: ['Primary email', 'EMAIL', 'Email address'] })
    ]
  },

  accounts: {
    sheet: 'MainAccounts', perm: 'masterdata', key: ['MAINACCOUNTID'], title: { ar: 'دليل الحسابات', en: 'Chart of accounts' },
    sections: [{ id: 'general', ar: 'عام', en: 'General', groups: [
      { id: 'id', ar: 'الحساب', en: 'Account' }, { id: 'cls', ar: 'التصنيف', en: 'Classification' }] }],
    fields: [
      F('MAINACCOUNTID', 'رقم الحساب', 'Account number', 'text', { req: 1, key: 1, sec: 'general', grp: 'id', list: 1,
        aliases: ['Main account', 'Main account ID', 'MAINACCOUNT', 'Account', 'AccountNumber'] }),
      F('NAME', 'اسم الحساب', 'Account name', 'text', { req: 1, sec: 'general', grp: 'id', list: 1, aliases: ['Name', 'AccountName'] }),
      F('MAINACCOUNTCATEGORY', 'مجموعة الحساب الرئيسي', 'Main account group', 'text', { sec: 'general', grp: 'cls', list: 1,
        aliases: ['Main account category', 'MainAccountGroup', 'Account category', 'MAINACCOUNTCATEGORYID'] }),
      F('MAINACCOUNTTYPE', 'نوع الحساب', 'Account type', 'select', { opts: ACCOUNT_TYPES, req: 1, sec: 'general', grp: 'cls', list: 1,
        aliases: ['Main account type', 'Type', 'AccountType'] })
    ]
  },

  lookups: {
    sheet: 'Lookups', perm: 'masterdata', key: ['TYPE', 'CODE'], title: { ar: 'القوائم المرجعية', en: 'Lookups' },
    sections: [{ id: 'general', ar: 'عام', en: 'General', groups: [{ id: 'g', ar: 'القيمة', en: 'Value' }] }],
    fields: [
      F('TYPE', 'النوع', 'Type', 'select', { opts: LOOKUP_TYPES, req: 1, key: 1, sec: 'general', grp: 'g', list: 1 }),
      F('CODE', 'الكود', 'Code', 'text', { req: 1, key: 1, sec: 'general', grp: 'g', list: 1, aliases: ['Value', 'ID'] }),
      F('NAME', 'الاسم', 'Name', 'text', { sec: 'general', grp: 'g', list: 1, aliases: ['Description'] }),
      F('PARENT', 'الكود الأب', 'Parent code', 'lookup', { lkBy: 'TYPE', lkMap: LOOKUP_PARENTS, sec: 'general', grp: 'g', list: 1 })
    ]
  }
};

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

    const lt = readTable_('Lookups');
    if (!lt.rows.some(function (r) { return r.CODE; })) {
      [
        ['AssetGroup', 'BLD', 'المباني', ''], ['AssetGroup', 'VEH', 'السيارات', ''], ['AssetGroup', 'MED', 'الأجهزة الطبية', ''],
        ['AssetGroup', 'FUR', 'الأثاث', ''], ['AssetGroup', 'IT', 'أجهزة الحاسب', ''],
        ['OperatingUnit', 'HE01', 'مستشفيات', ''],
        ['Sub1_OperatingUnit', 'DJI_HE0101', 'مستشفى جيبوتي', 'HE01'],
        ['Sub2_OperatingUnit', 'DJI_HE010001', 'عام المستشفى', 'DJI_HE0101'],
        ['JournalName', 'FAJ', 'Fixed asset journal - يومية الأصول', ''],
        ['JournalName', 'FADEP', 'Depreciation - الإهلاك', ''],
        ['Location', 'MAIN', 'المبنى الرئيسي', ''], ['Department', 'ADMIN', 'الإدارة', '']
      ].forEach(function (a) {
        appendRow_(lt, { TYPE: a[0], CODE: a[1], NAME: a[2], PARENT: a[3], CREATEDBY: 'setup', CREATEDON: stamp_(), MODIFIEDBY: '', MODIFIEDON: '' });
      });
    }
    // المرحلة 2: أمثلة للمباني والغرف والموظفين إن لم تكن موجودة
    const types = {};
    lt.rows.forEach(function (r) { types[r.TYPE] = 1; });
    [['Building', 'B01', 'المبنى الإداري', ''], ['Room', 'R101', 'غرفة 101', 'B01'], ['Employee', 'E001', 'أمين العهدة', '']]
      .forEach(function (a) {
        if (types[a[0]]) return;
        appendRow_(lt, { TYPE: a[0], CODE: a[1], NAME: a[2], PARENT: a[3], CREATEDBY: 'setup', CREATEDON: stamp_(), MODIFIEDBY: '', MODIFIEDON: '' });
      });
    // نقل الحسابات الرئيسية القديمة (إن وُجدت) إلى دليل الحسابات
    const at = readTable_('MainAccounts');
    if (!at.rows.some(function (r) { return r.MAINACCOUNTID; })) {
      lt.rows.filter(function (r) { return r.TYPE === 'MainAccount' && r.CODE; }).forEach(function (r) {
        appendRow_(at, { MAINACCOUNTID: r.CODE, NAME: r.NAME, MAINACCOUNTCATEGORY: '', MAINACCOUNTTYPE: 'BalanceSheet', CREATEDBY: 'setup', CREATEDON: stamp_() });
      });
    }

    const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('ورقة1');
    if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
    PropertiesService.getScriptProperties().setProperty('FA_SETUP', SETUP_VERSION);
  });
  CacheService.getScriptCache().remove('SETTINGS');
}

/** بيانات تجريبية اختيارية مطابقة للصور (BLD_00001). */
function seedDemoData() {
  ensureSetup_();
  withLock_(function () {
    const at = readTable_('FixedAssets');
    if (!at.rows.some(function (r) { return r.FIXEDASSETNUMBER === 'BLD_00001'; })) {
      appendRow_(at, { FIXEDASSETNUMBER: 'BLD_00001', FIXEDASSETGROUPID: 'BLD', NAME: 'غرفة الومتال +تشطيب السقف', ASSETTYPE: 'Tangible',
        QUANTITY: 1, MODEL: '', GUARANTEEDATE: '', LOCATION: 'MAIN', DEPARTMENT: 'ADMIN', ROOMNUMBER: '', RESPONSIBLE: '',
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
  else if (role === 'Accountant') { set(data, ['view', 'add', 'edit']); set(['depreciation'], ['view', 'add']); set(['reports'], ['view']); }
  else if (role !== 'none') set(data.concat(['reports']), ['view']);
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
    clientSchemas[e] = { perm: x.perm, key: x.key, title: x.title, sections: x.sections, fields: x.fields, autoId: x.autoId || '' };
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
    lookupTypes: LOOKUP_TYPES,
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
  LOOKUP_TYPES.forEach(function (t) { out[t.v] = []; });
  readTable_('Lookups').rows.forEach(function (r) {
    if (!r.TYPE || !r.CODE) return;
    (out[r.TYPE] = out[r.TYPE] || []).push({ c: String(r.CODE), n: String(r.NAME || ''), p: String(r.PARENT || '') });
  });
  Object.keys(out).forEach(function (k) { out[k].sort(function (a, b) { return a.c.localeCompare(b.c); }); });
  return out;
}

function refs_(s) {
  const out = { assets: [], vendors: [], books: [], accounts: [] };
  if (['journal', 'masterdata', 'depreciation', 'settings'].some(function (pg) { return can_(s, pg, 'view'); })) {
    out.accounts = readTable_('MainAccounts').rows.filter(function (r) { return r.MAINACCOUNTID; })
      .map(function (r) { return { c: String(r.MAINACCOUNTID), n: String(r.NAME || '') }; });
  }
  if (['assets', 'books', 'journal', 'depreciation', 'reports'].some(function (pg) { return can_(s, pg, 'view'); })) {
    out.assets = readTable_('FixedAssets').rows.filter(function (r) { return r.FIXEDASSETNUMBER; })
      .map(function (r) { return { c: String(r.FIXEDASSETNUMBER), n: String(r.NAME || '') }; });
    out.books = readTable_('AssetBooks').rows.filter(function (r) { return r.FIXEDASSETNUMBER; })
      .map(function (r) {
        return { a: String(r.FIXEDASSETNUMBER), c: String(r.BOOKID), n: String(r.STATUS || ''),
          ou: r.OPERATINGUNIT, s1: r.SUB1_OPERATINGUNIT, s2: r.SUB2_OPERATINGUNIT };
      });
  }
  if (can_(s, 'vendors', 'view') || can_(s, 'journal', 'view')) {
    out.vendors = readTable_('Vendors').rows.filter(function (r) { return r.VENDORACCOUNTNUMBER; })
      .map(function (r) { return { c: String(r.VENDORACCOUNTNUMBER), n: String(r.VENDORORGANIZATIONNAME || '') }; });
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
    asset: function (n) {
      const i = ctx.index('assets')[low_(n)];
      return i === undefined ? null : ctx.table('assets').rows[i];
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
    batch: function () { if (!ctx.importBatch) ctx.importBatch = nextBatch_(); return ctx.importBatch; }
  };
  return ctx;
}

function nextBatch_() {
  const st = getSettings_();
  const n = Number(st.JOURNALSEQ || 0) + 1;
  setSettings_({ JOURNALSEQ: n });
  return (st.JOURNALPREFIX || 'FAJ-') + ('000000' + n).slice(-6);
}

const HOOKS = {
  assets: function (rec) {
    if (rec.QUANTITY !== '' && rec.QUANTITY < 0) throw new Error('الكمية لا يمكن أن تكون سالبة');
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
    if (!rec.OPERATINGUNIT && rec.DEFAULTDIMENSIONDISPLAYVALUE) {
      const p = dimSplit_(rec.DEFAULTDIMENSIONDISPLAYVALUE);
      rec.OPERATINGUNIT = p[0] || ''; rec.SUB1_OPERATINGUNIT = p[1] || ''; rec.SUB2_OPERATINGUNIT = p[2] || '';
    }
    rec.DEFAULTDIMENSIONDISPLAYVALUE = dimJoin_([rec.OPERATINGUNIT, rec.SUB1_OPERATINGUNIT, rec.SUB2_OPERATINGUNIT]);
    if (rec.ACCUMULATEDDEPRECIATION !== '' && rec.ACCUMULATEDDEPRECIATION < 0) rec.ACCUMULATEDDEPRECIATION = -rec.ACCUMULATEDDEPRECIATION; // D365 يصدّره بالسالب
    if (num_(rec.ACCUMULATEDDEPRECIATION) > num_(rec.ACQUISITIONPRICE)) throw new Error('مجمع الإهلاك أكبر من سعر الاقتناء');
    rec.NETBOOKVALUE = nbv_(rec);
  },

  lookups: function (rec, ex, ctx) {
    const pt = LOOKUP_PARENTS[rec.TYPE];
    if (!pt) { rec.PARENT = ''; return; }
    if (rec.PARENT && !ctx.index('lookups')[low_(pt) + '|' + low_(rec.PARENT)]) throw new Error('الكود الأب غير موجود: ' + rec.PARENT);
  },

  journal: function (rec, ex, ctx) {
    if (ex && ex.STATUS && ex.STATUS !== 'Draft') throw new Error('لا يمكن تعديل سطر معتمد أو مُرحّل');
    if (ex) rec.JOURNALBATCHNUMBER = ex.JOURNALBATCHNUMBER;
    if (!rec.JOURNALBATCHNUMBER) rec.JOURNALBATCHNUMBER = ctx.batch();
    if (!ex) {
      const st = ctx.batchStatus(rec.JOURNALBATCHNUMBER);
      if (st && st !== 'Draft') throw new Error('لا يمكن إضافة أسطر إلى يومية معتمدة أو مُرحّلة: ' + rec.JOURNALBATCHNUMBER);
    }
    const a = ctx.asset(rec.FIXEDASSETNUMBER);
    if (!a) throw new Error('رقم الأصل غير موجود: ' + rec.FIXEDASSETNUMBER);
    rec.FIXEDASSETNUMBER = a.FIXEDASSETNUMBER;
    rec.ASSETNAME = a.NAME;
    if (!rec.BOOKID) rec.BOOKID = 'General';
    const d = num_(rec.DEBITAMOUNT), c = num_(rec.CREDITAMOUNT);
    if (d < 0 || c < 0) throw new Error('المبالغ لا يمكن أن تكون سالبة');
    if (d > 0 && c > 0) throw new Error('أدخل مبلغاً في المدين أو الدائن فقط، وليس كليهما');
    if (!d && !c) throw new Error('أدخل مبلغ المدين أو الدائن');
    rec.DEBITAMOUNT = d; rec.CREDITAMOUNT = c;
    if (!rec.OFFSETACCOUNTTYPE) rec.OFFSETACCOUNTTYPE = 'Ledger';
    if (!rec.OFFSETACCOUNT && rec.OFFSETACCOUNTDISPLAYVALUE) {
      const p = dimSplit_(rec.OFFSETACCOUNTDISPLAYVALUE);
      rec.OFFSETACCOUNT = p[0] || '';
      if (rec.OFFSETACCOUNTTYPE === 'Ledger') {
        rec.OFFSET_OPERATINGUNIT = p[1] || ''; rec.OFFSET_SUB1_OPERATINGUNIT = p[2] || ''; rec.OFFSET_SUB2_OPERATINGUNIT = p[3] || '';
      }
    }
    if (rec.OFFSETACCOUNTTYPE === 'Vendor') {
      rec.OFFSET_OPERATINGUNIT = rec.OFFSET_SUB1_OPERATINGUNIT = rec.OFFSET_SUB2_OPERATINGUNIT = '';
      rec.OFFSETACCOUNTDISPLAYVALUE = rec.OFFSETACCOUNT;
    } else {
      rec.OFFSETACCOUNTDISPLAYVALUE = rec.OFFSETACCOUNT ? dimJoin_([rec.OFFSETACCOUNT, rec.OFFSET_OPERATINGUNIT,
        rec.OFFSET_SUB1_OPERATINGUNIT, rec.OFFSET_SUB2_OPERATINGUNIT]) : '';
    }
    if (rec.TRANSACTIONTYPE === 'Depreciation' && rec.DEPRECIATIONMONTHS === '') rec.DEPRECIATIONMONTHS = 1;
    if (!rec.STATUS) rec.STATUS = 'Draft';
  }
};

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
  if (HOOKS[entity]) HOOKS[entity](rec, ex, ctx);
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
      if (entity === 'journal' && r.STATUS !== 'Draft') throw new Error('لا يمكن حذف سطر معتمد أو مُرحّل');
      if (entity === 'assets') {
        const n = low_(r.FIXEDASSETNUMBER);
        if (ctx.table('books').rows.some(function (b) { return low_(b.FIXEDASSETNUMBER) === n; }) ||
            ctx.table('journal').rows.some(function (j) { return low_(j.FIXEDASSETNUMBER) === n; })) {
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
        if (ctx.table('journal').rows.some(function (j) { return j.OFFSETACCOUNTTYPE === 'Vendor' && low_(j.OFFSETACCOUNT) === v; })) {
          throw new Error('لا يمكن حذف المورد ' + r.VENDORACCOUNTNUMBER + ' لوجود معاملات مرتبطة به');
        }
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
    rows.forEach(function (raw, n) {
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
      postToBooks_(lines, ctx);
      lines.forEach(function (l) { l.STATUS = 'Posted'; l.POSTEDBY = s.username; l.POSTEDON = now; });
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
        if (!b.LASTDEPRECIATIONDATE || String(l.TRANSDATE) > String(b.LASTDEPRECIATIONDATE)) b.LASTDEPRECIATIONDATE = l.TRANSDATE;
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
 * يحسب مقترح الإهلاك (القسط الثابت حسب العمر الافتراضي):
 *   الإهلاك = سعر الاقتناء ÷ العمر بالسنوات ÷ 12 × عدد الشهور
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
    if (fOU && low_(b.OPERATINGUNIT) !== fOU) return;
    const id = b.FIXEDASSETNUMBER;
    const skip = function (why) { skipped.push({ asset: id, name: a.NAME || '', book: b.BOOKID, reason: why }); };
    if (b.STATUS !== 'Open') return skip('حالة الدفتر ليست مفتوح');
    const price = num_(b.ACQUISITIONPRICE), life = num_(b.SERVICELIFE);
    if (price <= 0) return skip('لا يوجد سعر اقتناء');
    if (life <= 0) return skip('لا يوجد عمر افتراضي');
    if (pending[low_(id) + '|' + low_(b.BOOKID)]) return skip('يوجد إهلاك غير مُرحّل في اليومية ' + pending[low_(id) + '|' + low_(b.BOOKID)]);
    const startYM = b.LASTDEPRECIATIONDATE ? ym_(b.LASTDEPRECIATIONDATE) + 1 : ym_(b.DEPRECIATIONSTARTDATE || b.ACQUISITIONDATE);
    if (startYM === null) return skip('لا يوجد تاريخ بداية للإهلاك');
    let n = endYM - startYM + 1;
    if (n <= 0) return skip('تم إهلاكه حتى ' + b.LASTDEPRECIATIONDATE);
    if (b.DEPRECIATIONPERIODSREMAINING !== '' && b.DEPRECIATIONPERIODSREMAINING !== null) n = Math.min(n, Math.floor(num_(b.DEPRECIATIONPERIODSREMAINING)));
    if (n <= 0) return skip('لا توجد فترات متبقية');
    const accum = num_(b.ACCUMULATEDDEPRECIATION);
    const room = round_(price - num_(b.SCRAPVALUE) - accum, 3);
    if (room <= 0) return skip('وصل إلى قيمة الخردة');
    const monthly = price / life / 12;
    const amount = Math.min(round_(monthly * n, 3), room);
    lines.push({
      FIXEDASSETNUMBER: id, NAME: a.NAME || '', GROUP: a.FIXEDASSETGROUPID || '', BOOKID: b.BOOKID,
      PRICE: price, LIFE: life, ANNUAL: round_(price / life, 3), FROM: ymStr_(startYM), TO: end, MONTHS: n,
      CATCHUP: n > months, AMOUNT: -amount, ACCUM: accum, ACCUMAFTER: round_(accum + amount, 3), NBVAFTER: round_(price - accum - amount, 3),
      OU: b.OPERATINGUNIT || '', SUB1: b.SUB1_OPERATINGUNIT || '', SUB2: b.SUB2_OPERATINGUNIT || ''
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
        DEBITAMOUNT: 0, CREDITAMOUNT: -l.AMOUNT, DEPRECIATIONMONTHS: l.MONTHS, OFFSETACCOUNTTYPE: 'Ledger', OFFSETACCOUNT: acct,
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
      OPERATINGUNIT: b.OPERATINGUNIT || '', SUB1_OPERATINGUNIT: b.SUB1_OPERATINGUNIT || '', SUB2_OPERATINGUNIT: b.SUB2_OPERATINGUNIT || ''
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
