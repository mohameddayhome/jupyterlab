/**
 * المرحلة الرابعة: التقارير
 *  - تقرير أداء المشروعات: المستلم / المنصرف / بنود الصرف / المتبقي (لفترة اختيارية)
 *  - تقرير نسبة التنفيذ: المدة المنقضية مقابل التنفيذ المالي والإنجاز الفعلي، والمتبقي من المخصص (حتى تاريخ معيّن)
 * كل الأرقام من السندات المرحّلة فقط وبنفس طريقة التحويل المستخدمة في الأرصدة وكشف الحساب.
 */

var EXEC_STATUS = {
  notplanned: 'تواريخ غير محددة',
  notstarted: 'لم يبدأ',
  ontrack: 'في الموعد',
  late: 'متأخر',
  expired: 'انتهت المدة',
  done: 'مكتمل',
  over: 'تجاوز المخصص'
};

function projectBase_(p) {
  return {
    ProjectId: String(p.ProjectId), Description: p.Description, ProjectGroup: p.ProjectGroup, ProjectType: p.ProjectType,
    AccountId: String(p.AccountId), StartDate: dateKey_(p.ProjectedStartDate), EndDate: dateKey_(p.ProjectedEndDate),
    Budget: p.Budget === '' || p.Budget === null ? null : Number(p.Budget) || 0,
    BudgetCurrency: p.BudgetCurrency || 'KWD',
    CompletionPct: p.CompletionPct === '' || p.CompletionPct === null ? null : Number(p.CompletionPct)
  };
}

function apiPerformance(token, fromDate, toDate) {
  return run_(function () {
    var u = requireUser_(token);
    requirePerm_(u, 'performance', 'view');
    var from = dateKey_(fromDate), to = dateKey_(toDate);
    if (from && to && from > to) throw new Error('تاريخ البداية بعد تاريخ النهاية.');
    var agg = projectAgg_({ from: from, to: to, items: true });
    return {
      from: from, to: to,
      rows: readTable_('Projects').map(function (p) {
        var r = projectBase_(p);
        var x = agg[r.ProjectId];
        ['KWD', 'USD', 'DJF'].forEach(function (c) {
          r['in' + c] = x ? x['in' + c] : 0;
          r['out' + c] = x ? x['out' + c] : 0;
          r['bal' + c] = x ? x['bal' + c] : 0;
        });
        r.items = x ? x.items : [];
        return r;
      })
    };
  });
}

function daysBetween_(a, b) {
  var d1 = new Date(a + 'T00:00:00Z'), d2 = new Date(b + 'T00:00:00Z');
  return Math.round((d2 - d1) / 86400000);
}

function apiExecution(token, asOfDate) {
  return run_(function () {
    var u = requireUser_(token);
    requirePerm_(u, 'execution', 'view');
    var asOf = dateKey_(asOfDate) || Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
    var agg = projectAgg_({ to: asOf });
    var rows = readTable_('Projects').map(function (p) {
      var r = projectBase_(p);
      var x = agg[r.ProjectId];
      var c = r.BudgetCurrency;
      var dec = { KWD: 3, USD: 2, DJF: 0 }[c];
      r.received = x ? x['in' + c] : 0;
      r.spent = x ? x['out' + c] : 0;
      r.cash = round_(r.received - r.spent, dec);

      // نسبة المدة المنقضية
      r.durationDays = null; r.elapsedDays = null; r.elapsedPct = null; r.remainingDays = null;
      if (r.StartDate && r.EndDate) {
        r.durationDays = Math.max(1, daysBetween_(r.StartDate, r.EndDate));
        r.elapsedDays = Math.min(Math.max(0, daysBetween_(r.StartDate, asOf)), r.durationDays);
        r.elapsedPct = round_(r.elapsedDays / r.durationDays * 100, 1);
        r.remainingDays = Math.max(0, daysBetween_(asOf, r.EndDate));
      }
      // التنفيذ المالي والمتبقي من المخصص
      r.finPct = r.Budget ? round_(r.spent / r.Budget * 100, 1) : null;
      r.fundedPct = r.Budget ? round_(r.received / r.Budget * 100, 1) : null;
      r.remainingBudget = r.Budget !== null ? round_(r.Budget - r.spent, dec) : null;
      r.progressPct = r.CompletionPct !== null ? r.CompletionPct : r.finPct;

      // الحالة
      var s;
      if (r.Budget && r.spent > r.Budget) s = 'over';
      else if (!r.StartDate || !r.EndDate) s = 'notplanned';
      else if (asOf < r.StartDate) s = 'notstarted';
      else if ((r.progressPct || 0) >= 100) s = 'done';
      else if (asOf > r.EndDate) s = 'expired';
      else if (r.elapsedPct - (r.progressPct || 0) > 20) s = 'late';
      else s = 'ontrack';
      r.status = s; r.statusLabel = EXEC_STATUS[s];
      r.gap = r.elapsedPct !== null && r.progressPct !== null ? round_(r.elapsedPct - r.progressPct, 1) : null;
      return r;
    });
    return { asOf: asOf, statuses: EXEC_STATUS, rows: rows };
  });
}
