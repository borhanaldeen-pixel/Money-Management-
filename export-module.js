/* =========================================================
   export-module.js — تصدير Excel و PDF لصفحات: الراتب، الديون، الرادار المالي
   يُحمَّل بعد السكربت الرئيسي مباشرة، ويحقن أزرار التصدير بنفسه.
========================================================= */
(function () {
'use strict';

const LIBS = {
  xlsx: 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js',
  pdf: 'https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js'
};
const loaded = {};
const loadLib = k => loaded[k] || (loaded[k] = new Promise((ok, no) => {
  const s = document.createElement('script');
  s.src = LIBS[k]; s.onload = ok;
  s.onerror = () => { delete loaded[k]; no(new Error('lib')); };
  document.head.appendChild(s);
}));

const A = h => 'FF' + h;
const sum = (rows, j) => rows.reduce((a, r) => a + (+r[j] || 0), 0);
const fnum = n => Number(n || 0).toLocaleString('en-US', { maximumFractionDigits: 2 });
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' }[m]));
const today = () => getTodayISO(0);
const isChar = e => typeof e.catId === 'string' && e.catId.indexOf('sadaqah') === 0;
const GREEN = { bg: 'D9F2E3', fg: '187B30' }, RED = { bg: 'FBE0E8', fg: 'C01C57' }, GOLD = { bg: 'FFF2C3', fg: '8A5A10' };

/* ---------------- نماذج البيانات (تُستخدم في Excel و PDF معاً) ---------------- */
function salaryM() {
  const g = id => Math.max(0, parseFloat($(id).value) || 0);
  const sal = g('salaryInput'), h50 = g('hours50Input'), h100 = g('hours100Input'), rest = g('restaurantDeductionInput'), oth = g('otherDeductionInput');
  const rate = sal / 240, o50 = h50 * rate * 1.5, o100 = h100 * rate * 2, gross = sal + o50 + o100;
  const base = tax15Mode === 'both' ? gross : tax15Mode === 'overtime' ? o50 + o100 : sal;
  const tax = base * 0.15, s7 = sal * 0.07, u6 = sal * 0.006, ded = tax + s7 + u6 + rest + oth;
  const rows = [
    ['مستحقات', 'الراتب الشامل', 'الراتب الأساسي المعتمد', sal],
    ['مستحقات', 'إضافي 50%', fnum(h50) + ' ساعة × ' + fnum(rate * 1.5) + ' ر.ي', o50],
    ['مستحقات', 'إضافي 100%', fnum(h100) + ' ساعة × ' + fnum(rate * 2) + ' ر.ي', o100],
    ['خصومات', 'ضريبة 15%', tax15Mode === 'both' ? 'من الراتب + الإضافي' : tax15Mode === 'overtime' ? 'من الإضافي فقط' : 'من الراتب فقط', -tax],
    ['خصومات', 'تأمين اجتماعي 7%', 'من الراتب الشامل', -s7],
    ['خصومات', 'تأمين المتحدة 0.6%', 'من الراتب الشامل', -u6],
    ['خصومات', 'خصميات المطعم', 'إدخال يدوي', -rest],
    ['خصومات', 'خصميات أخرى', 'إدخال يدوي', -oth]
  ];
  rows.forEach(r => r.push(gross > 0 ? Math.abs(r[3]) / gross : 0));
  return { sheets: [{
    name: 'كشف الراتب', title: 'كشف احتساب الراتب المستلم', tab: 'F12761',
    kpis: [{ l: 'إجمالي المستحقات', v: gross, c: '187B30' }, { l: 'إجمالي الخصومات', v: ded, c: 'C01C57' },
           { l: 'صافي الراتب المستلم', v: gross - ded, c: '005245' }, { l: 'قيمة الساعة العادية', v: rate, c: '740E4C' }],
    cols: [{ h: 'القسم', w: 14, f: 'c', cf: v => v === 'مستحقات' ? GREEN : RED }, { h: 'البند', w: 26, f: 't' }, { h: 'البيان', w: 34, f: 't' },
           { h: 'المبلغ (ر.ي)', w: 18, f: 'n' }, { h: 'النسبة من المستحقات', w: 16, f: 'p' }],
    rows, sum: [3], sumLabel: 'صافي الراتب المستلم'
  }] };
}

function debtsM() {
  const list = debts.map(d => { const paid = d.paid || 0, rem = Math.max(0, d.amount - paid); return { d, paid, rem, cl: d.closed || rem === 0 }; })
    .sort((a, b) => a.cl - b.cl || b.rem - a.rem);
  const rows = list.map((x, i) => [i + 1, x.d.name, x.d.amount, x.paid, x.rem, x.d.amount > 0 ? x.paid / x.d.amount : 1,
    x.cl ? 'مسدد بالكامل' : x.paid > 0 ? 'سداد جزئي' : 'نشط', x.d.date || '', x.d.dueDate || '']);
  const hist = [];
  debts.forEach(d => (d.history && d.history.length ? d.history : [{ type: 'add', amount: d.amount, date: d.date, note: 'رصيد ابتدائي' }]).forEach(h =>
    hist.push([h.date || d.date || '', d.name, h.type === 'payment' ? 'سداد' : 'قيد دين', h.amount, h.note || '', h.type === 'payment' ? (h.deductedFromSalary === false ? 'لا' : 'نعم') : '—'])));
  hist.sort((a, b) => String(b[0]).localeCompare(String(a[0])));
  const active = list.filter(x => !x.cl).length;
  return { sheets: [{
    name: 'دفتر الديون', title: 'دفتر الديون والالتزامات', tab: '740E4C',
    kpis: [{ l: 'إجمالي الديون النشطة', v: computedActiveDebtTotal, c: '53051D' }, { l: 'المسدد منها', v: computedActiveDebtPaid, c: '187B30' },
           { l: 'صافي الدين المتبقي', v: computedActiveDebtRemaining, c: 'C01C57' }, { l: 'عدد الحسابات النشطة', v: active, c: '740E4C', q: 1 }],
    cols: [{ h: '#', w: 6, f: 'c' }, { h: 'الجهة / الدائن', w: 26, f: 't' }, { h: 'إجمالي الدين', w: 16, f: 'n' }, { h: 'المسدد', w: 16, f: 'n' },
           { h: 'المتبقي', w: 16, f: 'n', fx: r => 'C' + r + '-D' + r }, { h: 'نسبة السداد', w: 13, f: 'p', fx: r => 'IF(C' + r + '=0,1,D' + r + '/C' + r + ')' },
           { h: 'الحالة', w: 16, f: 'c', cf: v => v === 'مسدد بالكامل' ? GREEN : v === 'نشط' ? RED : GOLD },
           { h: 'آخر حركة', w: 14, f: 'c' }, { h: 'ملاحظة', w: 28, f: 't' }],
    rows, sum: [2, 3, 4], sumLabel: 'الإجمالي',
    tot: { 5: { fx: r => 'IF(C' + r + '=0,0,D' + r + '/C' + r + ')', v: rs => { const c = sum(rs, 2); return c ? sum(rs, 3) / c : 0; } } }
  }, {
    name: 'كشف الحركات', title: 'كشف حركات الديون (قيود وسدادات)', tab: 'EEAC4D',
    cols: [{ h: 'التاريخ', w: 14, f: 'c' }, { h: 'الجهة', w: 26, f: 't' }, { h: 'النوع', w: 14, f: 'c', cf: v => v === 'سداد' ? GREEN : RED },
           { h: 'المبلغ (ر.ي)', w: 16, f: 'n' }, { h: 'الملاحظة', w: 34, f: 't' }, { h: 'مخصوم من الراتب', w: 16, f: 'c' }],
    rows: hist
  }] };
}

function radarM() {
  const ns = computedNetReceivedSalary, ch = getCharityTotals(), sav = Math.max(0, parseFloat($('savingsLockInput').value) || 0);
  const paidD = expensesList.filter(e => e.catId === 'debt_pay' || e.linkedDebtId).reduce((s, e) => s + e.amount, 0);
  const free = computedRemainingBalance - computedActiveDebtRemaining, P = v => ns > 0 ? v / ns : 0;
  const rows = [
    ['صافي الراتب المستلم', ns, 1, 'بعد الضريبة والتأمينات والخصميات'],
    ['الميزانية القابلة للصرف', computedSpendableBudget, P(computedSpendableBudget), 'بعد خزنة الادخار والرصيد الإضافي'],
    ['إجمالي المصروف', computedTotalSpent, P(computedTotalSpent), 'شامل سداد الديون والصدقة'],
    ['المتبقي من الراتب', computedRemainingBalance, P(computedRemainingBalance), 'السيولة القابلة للصرف الآن'],
    ['سداد الديون من الراتب', paidD, P(paidD), 'مخصوم تلقائياً من الراتب'],
    ['الصدقة اليومية', ch.total, P(ch.total), fnum(ch.pct) + '% من الراتب'],
    ['خزنة الادخار', sav, P(sav), 'محجوزة ولا تدخل في الصرف'],
    ['إجمالي الديون النشطة', computedActiveDebtTotal, P(computedActiveDebtTotal), ''],
    ['المتبقي من الديون', computedActiveDebtRemaining, P(computedActiveDebtRemaining), ''],
    ['الفائض الحقيقي بعد الديون', free, P(free), free >= 0 ? 'ملاءة مالية موجبة ✅' : 'السيولة لا تغطي الديون ⚠️']
  ];
  const sheets = [{
    name: 'الرادار المالي', title: 'الرادار المالي — الملاءة والتحليل', tab: '00ACA5',
    kpis: [{ l: 'المتبقي من الراتب', v: computedRemainingBalance, c: '005245' }, { l: 'الديون المتبقية', v: computedActiveDebtRemaining, c: 'C01C57' },
           { l: 'الفائض الحقيقي', v: free, c: free >= 0 ? '187B30' : 'C01C57' }, { l: 'الصدقة من الراتب', v: ch.pct, c: '8A5A10', p: 1 }],
    cols: [{ h: 'البند', w: 30, f: 't' }, { h: 'المبلغ (ر.ي)', w: 18, f: 'n' }, { h: 'من الراتب', w: 13, f: 'p' }, { h: 'ملاحظة', w: 38, f: 't' }],
    rows
  }];
  if (expensesList.length) {
    let bal = computedSpendableBudget;
    const log = [...expensesList].sort((a, b) => a.date !== b.date ? a.date.localeCompare(b.date) : a.timestamp - b.timestamp).map(e => {
      bal -= e.amount; const r = [e.date, e.catIcon + ' ' + e.catName, e.note || '', e.amount, bal]; r._c = e; return r; });
    sheets.push({
      name: 'سجل المصاريف', title: 'سجل المصاريف وتناقص الراتب', tab: 'F12761',
      cols: [{ h: 'التاريخ', w: 14, f: 'c' }, { h: 'الفئة', w: 24, f: 't', cf: (v, r) => isChar(r._c) ? GOLD : (r._c.catId === 'debt_pay' || r._c.linkedDebtId) ? RED : null },
             { h: 'الملاحظة', w: 34, f: 't' }, { h: 'المبلغ (ر.ي)', w: 16, f: 'n' },
             { h: 'المتبقي بعده', w: 18, f: 'n', fx: (r, h) => r === h + 1 ? computedSpendableBudget + '-D' + r : 'E' + (r - 1) + '-D' + r }],
      rows: log, sum: [3], sumLabel: 'إجمالي المصروف'
    });
    const by = {};
    expensesList.forEach(e => { const k = e.catIcon + ' ' + e.catName; by[k] = by[k] || [k, 0, 0, 0]; by[k][1] += e.amount; by[k][3]++; });
    const cr = Object.values(by).sort((a, b) => b[1] - a[1]), tt = sum(cr, 1);
    cr.forEach(r => r[2] = tt ? r[1] / tt : 0);
    sheets.push({
      name: 'الفئات', title: 'أين يذهب الراتب؟ (توزيع الفئات)', tab: 'EEAC4D',
      cols: [{ h: 'الفئة', w: 26, f: 't' }, { h: 'المبلغ (ر.ي)', w: 18, f: 'n' },
             { h: 'النسبة', w: 14, f: 'p', fx: (r, h) => 'IF(SUM(B$' + (h + 1) + ':B$' + (h + cr.length) + ')=0,0,B' + r + '/SUM(B$' + (h + 1) + ':B$' + (h + cr.length) + '))' },
             { h: 'عدد العمليات', w: 14, f: 'c' }],
      rows: cr, sum: [1, 3], sumLabel: 'الإجمالي'
    });
  }
  return { sheets };
}

function model(p) {
  if (p === 'all') return { file: 'التقرير_المالي_الشامل', title: 'التقرير المالي الشامل', sheets: [...salaryM().sheets, ...debtsM().sheets, ...radarM().sheets] };
  const t = { salary: ['كشف_الراتب', 'كشف الراتب والإضافي', salaryM], debts: ['دفتر_الديون', 'دفتر الديون والالتزامات', debtsM], hub: ['الرادار_المالي', 'الرادار المالي الشامل', radarM] }[p];
  return { file: t[0], title: t[1], sheets: t[2]().sheets };
}

/* ---------------- Excel (ExcelJS: تنسيق كامل) ---------------- */
function addSheet(wb, s, sub) {
  const n = s.cols.length, fill = h => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: A(h) } });
  const ws = wb.addWorksheet(s.name, { properties: { tabColor: { argb: A(s.tab || '740E4C') } },
    pageSetup: { paperSize: 9, orientation: n > 6 ? 'landscape' : 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, horizontalCentered: true } });
  ws.columns = s.cols.map(c => ({ width: c.w }));
  const bd = { style: 'thin', color: { argb: A('D9B7C6') } }, B = { top: bd, left: bd, bottom: bd, right: bd };
  const font = o => Object.assign({ name: 'Arial', size: 11 }, o);
  const al = h => ({ horizontal: h, vertical: 'middle', wrapText: true, readingOrder: 'rtl' });
  const NF = { n: '#,##0;[Red]-#,##0', p: '0%', t: 'General', c: 'General' };
  const band = (r, t, bg, fg, sz, ht) => {
    for (let c = 1; c <= n; c++) ws.getCell(r, c).fill = fill(bg);
    ws.mergeCells(r, 1, r, n);
    const x = ws.getCell(r, 1); x.value = t; x.font = font({ bold: true, size: sz, color: { argb: A(fg) } }); x.alignment = al('center');
    ws.getRow(r).height = ht;
  };
  band(1, s.title, '53051D', 'EEAC4D', 20, 40);
  band(2, sub, '740E4C', 'FFFFFF', 11, 22);

  let r = 4;
  if (s.kpis) {
    const k = s.kpis, sp = Math.max(1, Math.floor(n / k.length)); let c0 = 1;
    k.forEach((x, i) => {
      if (c0 > n) return;
      const c1 = i === k.length - 1 ? n : Math.min(n, c0 + sp - 1);
      for (let c = c0; c <= c1; c++) {
        ws.getCell(r, c).fill = fill('740E4C'); ws.getCell(r + 1, c).fill = fill('FFF3E1');
        ws.getCell(r, c).border = B; ws.getCell(r + 1, c).border = B;
      }
      if (c1 > c0) { ws.mergeCells(r, c0, r, c1); ws.mergeCells(r + 1, c0, r + 1, c1); }
      const a = ws.getCell(r, c0), b = ws.getCell(r + 1, c0);
      a.value = x.l; a.font = font({ bold: true, size: 10, color: { argb: A('FFF2C3') } }); a.alignment = al('center');
      b.value = x.v; b.numFmt = x.p ? '0.0"%"' : x.q ? '0' : '#,##0" ر.ي";[Red]-#,##0" ر.ي"';
      b.font = font({ bold: true, size: 16, color: { argb: A(x.c || '53051D') } }); b.alignment = al('center');
      c0 = c1 + 1;
    });
    ws.getRow(r).height = 22; ws.getRow(r + 1).height = 34; r += 3;
  }

  const hr = r;
  s.cols.forEach((c, i) => {
    const x = ws.getCell(hr, i + 1); x.value = c.h; x.fill = fill(s.hd || '005245');
    x.font = font({ bold: true, size: 12, color: { argb: A('FFFFFF') } }); x.alignment = al('center'); x.border = B;
  });
  ws.getRow(hr).height = 30;

  s.rows.forEach((row, i) => {
    const rr = hr + 1 + i, R = ws.getRow(rr); R.height = 22;
    s.cols.forEach((c, j) => {
      const x = R.getCell(j + 1), v = row[j];
      x.value = c.fx ? { formula: c.fx(rr, hr), result: v } : v;
      x.numFmt = NF[c.f || 't']; x.border = B; x.font = font({}); x.alignment = al(c.f === 't' ? 'right' : 'center');
      if (i % 2) x.fill = fill('F8EEF2');
      const m = c.cf && c.cf(v, row);
      if (m) { x.font = font({ bold: true, color: { argb: A(m.fg) } }); x.fill = fill(m.bg); }
    });
  });

  const last = hr + s.rows.length; let end = last;
  if (s.sum && s.rows.length) {
    const tr = last + 1, sj = Math.max(1, Math.min(...s.sum)); end = tr;
    s.cols.forEach((c, j) => {
      const x = ws.getCell(tr, j + 1); x.fill = fill('53051D'); x.border = B; x.alignment = al('center');
      x.font = font({ bold: true, size: 12, color: { argb: A('EEAC4D') } }); x.numFmt = NF[c.f || 't'];
      const L = ws.getColumn(j + 1).letter;
      if (j >= sj && s.sum.includes(j)) x.value = { formula: 'SUM(' + L + (hr + 1) + ':' + L + last + ')', result: sum(s.rows, j) };
      else if (j >= sj && s.tot && s.tot[j]) x.value = { formula: s.tot[j].fx(tr), result: s.tot[j].v(s.rows) };
    });
    ws.getCell(tr, 1).value = s.sumLabel || 'الإجمالي';
    if (sj > 1) ws.mergeCells(tr, 1, tr, sj);
    ws.getRow(tr).height = 26;
  }
  if (s.rows.length) ws.autoFilter = { from: { row: hr, column: 1 }, to: { row: last, column: n } };
  band(end + 2, '© المنظومة المالية المتكاملة • Bento Agate Pro', 'F6EEF1', '6E354B', 9, 20);
  ws.views = [{ rightToLeft: true, showGridLines: false, state: 'frozen', ySplit: hr, activeCell: 'A' + (hr + 1) }];
  ws.pageSetup.printTitlesRow = hr + ':' + hr;
  ws.headerFooter.oddFooter = '&C&P / &N';
}

async function xlBuild(sheets, sub) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'المنظومة المالية المتكاملة'; wb.created = new Date();
  sheets.forEach(s => addSheet(wb, s, sub));
  return new Blob([await wb.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

/* ---------------- PDF (يُرسم كصفحة HTML عربية ثم يُحوَّل) ---------------- */
function pdfHtml(m, sub) {
  const fv = (v, f) => f === 'n' ? fnum(v) : f === 'p' ? Math.round(v * 100) + '%' : esc(v);
  const kp = k => k.p ? fnum(k.v) + '%' : k.q ? k.v : fnum(k.v) + ' ر.ي';
  const td = 'padding:6px 5px;border:1px solid #D9B7C6;';
  const sec = s => {
    const sj = s.sum ? Math.max(1, Math.min(...s.sum)) : 0;
    const tot = s.sum && s.rows.length ? '<tr style="background:#53051D;color:#EEAC4D;font-weight:900"><td colspan="' + sj + '" style="' + td + 'text-align:center">' + esc(s.sumLabel || 'الإجمالي') + '</td>' +
      s.cols.slice(sj).map((c, i) => { const j = i + sj, v = s.sum.includes(j) ? sum(s.rows, j) : s.tot && s.tot[j] ? s.tot[j].v(s.rows) : '';
        return '<td style="' + td + 'text-align:center">' + (v === '' ? '' : fv(v, c.f)) + '</td>'; }).join('') + '</tr>' : '';
    return '<div style="margin-top:18px"><div style="font-weight:900;font-size:15px;color:#53051D;border-right:5px solid #EEAC4D;padding-right:9px;margin-bottom:9px">' + esc(s.title) + '</div>' +
      (s.kpis ? '<div style="display:flex;gap:8px;margin-bottom:11px;page-break-inside:avoid">' + s.kpis.map(k =>
        '<div style="flex:1;border-radius:12px;background:#FFF3E1;border-top:4px solid #' + (k.c || '740E4C') + ';padding:8px 4px;text-align:center"><div style="font-size:10px;color:#6E354B;font-weight:700">' + esc(k.l) +
        '</div><div style="font-size:14px;font-weight:900;color:#' + (k.c || '53051D') + ';margin-top:3px">' + kp(k) + '</div></div>').join('') + '</div>' : '') +
      '<table style="width:100%;border-collapse:collapse;font-size:11px"><thead><tr>' + s.cols.map(c => '<th style="' + td + 'background:#005245;color:#fff;text-align:center;font-weight:900">' + esc(c.h) + '</th>').join('') + '</tr></thead><tbody>' +
      s.rows.map((r, i) => '<tr style="page-break-inside:avoid;background:' + (i % 2 ? '#F8EEF2' : '#fff') + '">' + s.cols.map((c, j) => {
        const x = c.cf && c.cf(r[j], r);
        return '<td style="' + td + 'text-align:' + (c.f === 't' ? 'right' : 'center') + (x ? ';background:#' + x.bg + ';color:#' + x.fg + ';font-weight:800' : '') + '">' + fv(r[j], c.f) + '</td>'; }).join('') + '</tr>').join('') + tot + '</tbody></table></div>';
  };
  return '<div dir="rtl" style="width:720px;font-family:Cairo,Arial,sans-serif;color:#2B0816;background:#fff">' +
    '<div style="background:linear-gradient(135deg,#360212,#740E4C 62%,#005245);color:#fff;border-radius:16px;padding:18px 20px;border-bottom:4px solid #EEAC4D">' +
    '<div style="font-size:11px;color:#FFF2C3">💎 المنظومة المالية المتكاملة</div><div style="font-size:24px;font-weight:900;color:#EEAC4D;margin-top:4px">' + esc(m.title) +
    '</div><div style="font-size:11px;margin-top:5px;color:#fff">' + esc(sub) + '</div></div>' + m.sheets.map(sec).join('') +
    '<div style="margin-top:20px;text-align:center;font-size:10px;color:#6E354B;border-top:2px solid #EEAC4D;padding-top:8px">المنظومة المالية المتكاملة • Bento Agate Pro</div></div>';
}

async function pdfBuild(m, sub) {
  const host = document.createElement('div');
  host.style.cssText = 'position:fixed;left:-9999px;top:0;width:720px';
  host.innerHTML = pdfHtml(m, sub); document.body.appendChild(host);
  try {
    if (document.fonts && document.fonts.ready) await document.fonts.ready;
    return await html2pdf().set({ margin: [8, 8, 10, 8], image: { type: 'jpeg', quality: 0.95 },
      html2canvas: { scale: 2, useCORS: true, backgroundColor: '#ffffff' }, jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
      pagebreak: { mode: ['css', 'legacy'], avoid: 'tr' } }).from(host).outputPdf('blob');
  } finally { host.remove(); }
}

/* ---------------- التسليم: تحميل أو مشاركة مباشرة ---------------- */
let share = false;
async function deliver(blob, name) {
  const file = new File([blob], name, { type: blob.type });
  if (share && navigator.canShare && navigator.canShare({ files: [file] })) {
    try { await navigator.share({ files: [file], title: name }); return; } catch (e) { if (e.name === 'AbortError') return; }
  }
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}

async function run(p, k, btn) {
  if (btn.disabled) return;
  const old = btn.innerHTML; btn.disabled = true; btn.innerHTML = '⏳ جارٍ التجهيز...';
  try {
    const m = model(p), sub = 'تاريخ التقرير: ' + today(), name = m.file + '_' + today() + (k === 'xlsx' ? '.xlsx' : '.pdf');
    let blob;
    if (k === 'xlsx') { await loadLib('xlsx'); blob = await xlBuild(m.sheets, sub); }
    else { await loadLib('pdf'); blob = await pdfBuild(m, sub); }
    await deliver(blob, name);
    showToast(share ? '📤 تم تجهيز الملف للمشاركة' : '✅ تم إنشاء ' + (k === 'xlsx' ? 'ملف Excel' : 'ملف PDF') + ' بنجاح');
  } catch (err) {
    console.error(err); showToast('⚠️ تعذر التصدير، تأكد من الاتصال بالإنترنت لتحميل المكتبة أول مرة');
  }
  btn.disabled = false; btn.innerHTML = old;
}

/* ---------------- الواجهة: أزرار التصدير داخل كل صفحة ---------------- */
const CSS = '.xp-row{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:6px}' +
  '.xp-btn{border:1px solid rgba(255,255,255,.25);border-radius:15px;padding:13px 8px;font-weight:900;font-size:.86rem;color:#fff;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:6px;transition:transform .2s}' +
  '.xp-btn:hover{transform:translateY(-2px)}.xp-btn:disabled{opacity:.65;cursor:wait}' +
  '.xp-x{background:linear-gradient(135deg,var(--c22),var(--c21));box-shadow:0 8px 18px rgba(24,123,48,.35)}' +
  '.xp-p{background:linear-gradient(135deg,var(--c5),var(--burgundy-wine));box-shadow:0 8px 18px rgba(241,39,97,.3)}' +
  '.xp-a{background:linear-gradient(135deg,#b87b1e,var(--burgundy-dark));border-color:var(--c12);box-shadow:0 8px 18px rgba(238,172,77,.28)}' +
  '.xp-share{display:flex;align-items:center;gap:8px;margin-top:11px;font-size:.74rem;font-weight:800;color:var(--text-muted);cursor:pointer}' +
  '.xp-share input{width:17px;height:17px;accent-color:var(--c11)}';

function inject() {
  const st = document.createElement('style'); st.textContent = CSS; document.head.appendChild(st);
  [['pageSalary', 'salary', 'كشف الراتب والإضافي'], ['pageDebts', 'debts', 'دفتر الديون'], ['pageHub', 'hub', 'الرادار المالي']].forEach(([id, p, t]) => {
    const g = document.querySelector('#' + id + ' .bento-grid'); if (!g) return;
    const d = document.createElement('div'); d.className = 'glass-card bento-cell col-12';
    d.innerHTML = '<div class="cell-header"><h4 class="cell-title"><span class="dot-indicator" style="background:var(--c11);color:var(--c11)"></span>📤 تصدير ' + t + '</h4></div>' +
      '<p class="cell-subtitle">Excel منسّق بعناوين ملوّنة ومعادلات حيّة وفلاتر، أو تقرير PDF جاهز للطباعة والمشاركة</p><div class="xp-row">' +
      '<button type="button" class="xp-btn xp-x" data-xp="' + p + ':xlsx">📊 تصدير Excel</button><button type="button" class="xp-btn xp-p" data-xp="' + p + ':pdf">📄 تصدير PDF</button>' +
      (p === 'hub' ? '<button type="button" class="xp-btn xp-a" data-xp="all:xlsx">📚 تقرير شامل Excel</button><button type="button" class="xp-btn xp-a" data-xp="all:pdf">📚 تقرير شامل PDF</button>' : '') +
      '</div><label class="xp-share"><input type="checkbox" class="xp-sh"> 📤 مشاركة مباشرة بدل التحميل (واتساب وغيره)</label>';
    g.appendChild(d);
  });
  document.addEventListener('click', ev => { const b = ev.target.closest && ev.target.closest('[data-xp]'); if (b) { const a = b.dataset.xp.split(':'); run(a[0], a[1], b); } });
  document.addEventListener('change', ev => {
    if (ev.target.classList && ev.target.classList.contains('xp-sh')) { share = ev.target.checked; document.querySelectorAll('.xp-sh').forEach(c => c.checked = share); }
  });
}

window.XP = { model, xlBuild, pdfHtml };
if (typeof document !== 'undefined' && document.querySelector) inject();
})();
