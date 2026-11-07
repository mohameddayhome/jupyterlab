/**
 * المرحلة الثالثة: كشف حساب تفصيلي لكل مشروع (من السندات المرحّلة فقط).
 * الوارد: سطور سندات التحويل بالعملات الثلاث كما هي.
 * المنصرف: سطور سندات الصرف محوّلة للعملات الثلاث بمتوسط سعر تحويلات المشروع
 * (نفس طريقة تقرير أرصدة المشروعات حتى تتطابق الأرقام).
 */

function dateKey_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  var m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[0] : '';
}

/** تحويل مبلغ بعملة ما إلى العملات الثلاث بمعاملات المشروع */
function toEq_(amount, currency, b) {
  var o = { KWD: 0, USD: 0, DJF: 0 };
  amount = Number(amount) || 0;
  if (!b.rUSD || !b.rDJF) { if (o[currency] !== undefined) o[currency] = amount; return o; }
  var rate = { KWD: 1, USD: b.rUSD, DJF: b.rDJF }[currency];
  if (!rate) return o;
  var kwd = amount / rate;
  return { KWD: kwd, USD: kwd * b.rUSD, DJF: kwd * b.rDJF };
}

function apiStatement(token, projectId, fromDate, toDate) {
  return run_(function () {
    var u = requireUser_(token);
    requirePerm_(u, 'statement', 'view');
    projectId = String(projectId || '');
    if (!projectId) throw new Error('اختر المشروع.');
    var project = readTable_('Projects').filter(function (p) { return String(p.ProjectId) === projectId; })[0];
    if (!project) throw new Error('المشروع غير موجود.');
    var from = dateKey_(fromDate), to = dateKey_(toDate);
    if (from && to && from > to) throw new Error('تاريخ البداية بعد تاريخ النهاية.');

    var b = projectBalances_()[projectId] || emptyBalance_();
    var rows = [];

    // الوارد: سطور سندات التحويل المرحّلة
    var rc = {};
    readTable_('Receipts').forEach(function (h) { if (h.Status === 'Posted') rc[String(h.DocumentNumber)] = h; });
    readTable_('ReceiptLines').forEach(function (l) {
      var h = rc[String(l.DocumentNumber)];
      if (!h || String(l.ProjectId) !== projectId) return;
      rows.push({
        date: dateKey_(h.Date), entity: 'receipts', type: 'تحويل وارد', doc: String(h.DocumentNumber), line: Number(l.LineId) || 0,
        desc: l.Description || ('تحويل ' + (h.Bank ? 'عن طريق ' + h.Bank : '') + (h.Reference ? ' — مرجع ' + h.Reference : '')),
        account: String(l.AccountId) + ' - ' + (l.AccountName || ''),
        amount: Number(l.AmountKWD) || 0, currency: 'KWD',
        inKWD: Number(l.AmountKWD) || 0, inUSD: Number(l.AmountUSD) || 0, inDJF: Number(l.AmountDJF) || 0,
        outKWD: 0, outUSD: 0, outDJF: 0
      });
    });

    // المنصرف: سطور سندات الصرف المرحّلة
    var pv = {};
    readTable_('Payments').forEach(function (h) {
      if (h.Status === 'Posted' && String(h.ProjectId) === projectId) pv[String(h.VoucherNo)] = h;
    });
    readTable_('PaymentLines').forEach(function (l) {
      var h = pv[String(l.VoucherNo)];
      if (!h) return;
      var e = toEq_(l.Amount, l.Currency, b);
      rows.push({
        date: dateKey_(h.Date), entity: 'payments', type: 'صرف', doc: String(h.VoucherNo), line: Number(l.LineId) || 0,
        desc: l.Description ? (h.Description + ' — ' + l.Description) : h.Description,
        account: String(l.AccountId) + ' - ' + (l.AccountName || ''), item: l.ExpenseItemName || '',
        amount: Number(l.Amount) || 0, currency: l.Currency,
        inKWD: 0, inUSD: 0, inDJF: 0, outKWD: e.KWD, outUSD: e.USD, outDJF: e.DJF
      });
    });

    rows.sort(function (a, b2) {
      if (a.date !== b2.date) return a.date < b2.date ? -1 : 1;
      if (a.entity !== b2.entity) return a.entity === 'receipts' ? -1 : 1; // الوارد قبل الصرف في نفس اليوم
      if (a.doc !== b2.doc) return a.doc < b2.doc ? -1 : 1;
      return a.line - b2.line;
    });

    var C = [['KWD', 3], ['USD', 2], ['DJF', 0]];
    var opening = { KWD: 0, USD: 0, DJF: 0 };
    var inRange = [];
    rows.forEach(function (r) {
      if (from && r.date < from) { C.forEach(function (c) { opening[c[0]] += r['in' + c[0]] - r['out' + c[0]]; }); return; }
      if (to && r.date > to) return;
      inRange.push(r);
    });
    var run = { KWD: opening.KWD, USD: opening.USD, DJF: opening.DJF };
    var totals = { inKWD: 0, inUSD: 0, inDJF: 0, outKWD: 0, outUSD: 0, outDJF: 0 };
    inRange.forEach(function (r) {
      C.forEach(function (c) {
        var k = c[0], d = c[1];
        run[k] += r['in' + k] - r['out' + k];
        totals['in' + k] += r['in' + k]; totals['out' + k] += r['out' + k];
        r['in' + k] = round_(r['in' + k], d); r['out' + k] = round_(r['out' + k], d);
        r['bal' + k] = round_(run[k], d);
      });
    });
    var closing = {};
    C.forEach(function (c) {
      opening[c[0]] = round_(opening[c[0]], c[1]);
      closing[c[0]] = round_(run[c[0]], c[1]);
      totals['in' + c[0]] = round_(totals['in' + c[0]], c[1]);
      totals['out' + c[0]] = round_(totals['out' + c[0]], c[1]);
    });

    return {
      project: {
        id: projectId, name: project.Description, group: project.ProjectGroup, type: project.ProjectType,
        account: String(project.AccountId) + ' - ' + (project.AccountName || '')
      },
      rates: { rUSD: round_(b.rUSD, 6), rDJF: round_(b.rDJF / (b.rUSD || 1), 6) },
      from: from, to: to, opening: opening, closing: closing, totals: totals, rows: inRange
    };
  });
}
