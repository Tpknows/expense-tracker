/* Liquidation ZIP: kapareho ng format ng mga dating liquidation folder, hal. "OCT REIMBURSEMENT LIQUIDATION 2026_3":
   - {batch}.xlsx: LIQUIDATION tab (BUDGET / EXPENSE / REMAINING, tapos AMOUNT | REMARKS | DATE | REMARKS | RECEIPT # | NAME)
     at APPROVAL tab (sino at kailan nag-approve)
   - 1.jpg, 2.jpg, ... (ang numero = RECEIPT # sa Excel)
   - OB Travel Forms - {Month D, YYYY}.xlsx (kung may Travel/Transportation)
   Kailangan: XLSX (SheetJS), JSZip, OBGen (ob-form.js). */
(function (root) {
  'use strict';
  var TRAVEL = 'Travel/Transportation';
  var MONTHS_LONG = ['January','February','March','April','May','June','July','August','September','October','November','December'];

  function ymd(s) { var p = String(s).split('-').map(Number); return { y: p[0], m: p[1], d: p[2] }; }
  function serial(d) { return (Date.UTC(d.y, d.m - 1, d.d) - Date.UTC(1899, 11, 30)) / 86400000; }
  function longDate(dt) { return MONTHS_LONG[dt.getMonth()] + ' ' + dt.getDate() + ', ' + dt.getFullYear(); }

  // Pangalan ng batch: "{MON} {FUND} LIQUIDATION {YYYY}", dagdag na _2, _3 ... kapag may nauna na sa folder
  function batchName(monthLabel, fund, year, existingNames) {
    var base = monthLabel + ' ' + fund.toUpperCase() + ' LIQUIDATION ' + year;
    var norm = function (s) { return String(s).replace(/\.zip$/i, '').replace(/\s+/g, ' ').trim().toUpperCase(); };
    var taken = (existingNames || []).map(norm).filter(function (n) { return n === base || n.indexOf(base + '_') === 0; });
    if (!taken.length) return base;
    var n = 2;
    while (taken.indexOf(base + '_' + n) >= 0) n++;
    return base + '_' + n;
  }

  function liquidationBook(XLSX, opt, items) {
    var first = 8, last = first + items.length - 1;
    var aoa = [[], [], ['BUDGET', ''], ['EXPENSE', null], ['REMAINING', null], [],
      ['AMOUNT', 'REMARKS', 'DATE', 'REMARKS', 'RECEIPT #', 'NAME', 'EXPENSE ID']];
    items.forEach(function (it) {
      aoa.push([it.amount, it.desc, serial(ymd(it.date)), it.particular, it.receiptNo, String(it.name).toUpperCase(), it.expenseId]);
    });
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws.B4 = { t: 'n', f: 'SUM(A' + first + ':A' + last + ')' };
    ws.B5 = { t: 'n', f: 'IF(B3="","",B3-B4)' };
    for (var r = first; r <= last; r++) {
      if (ws['A' + r]) ws['A' + r].z = '#,##0.00';
      if (ws['C' + r]) ws['C' + r].z = 'm/d/yyyy';
    }
    ws['!cols'] = [{ wch: 12 }, { wch: 48 }, { wch: 12 }, { wch: 22 }, { wch: 11 }, { wch: 24 }, { wch: 22 }];

    var ap = XLSX.utils.aoa_to_sheet([
      ['Approved by', opt.approver.name],
      ['Email', opt.approver.email],
      ['Date approved', longDate(opt.approvedAt)],
      ['Fund', opt.fund],
      ['Batch', opt.batch],
      ['Bilang ng resibo', items.length],
      ['Total', null],
      [],
      ['In-approve sa EasyPC Expense Log app.']
    ]);
    ap.B7 = { t: 'n', f: "LIQUIDATION!B4", z: '#,##0.00' };
    ap['!cols'] = [{ wch: 18 }, { wch: 44 }];

    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'LIQUIDATION');
    XLSX.utils.book_append_sheet(wb, ap, 'APPROVAL');
    return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  }

  /* opt: {fund, batch, approver:{name,email}, approvedAt:Date, people:{NAMEKEY:{name,id,department,superior,head}}}
     items: [{date:'YYYY-MM-DD', amount, desc, particular, name, expenseId, from, to, ride, receipt:{data, ext}}] */
  function build(libs, opt, items) {
    var XLSX = libs.XLSX, JSZip = libs.JSZip, OBGen = libs.OBGen;
    var sorted = items.slice().sort(function (a, b) {
      return a.date < b.date ? -1 : a.date > b.date ? 1 : String(a.name).localeCompare(String(b.name));
    });
    sorted.forEach(function (it, i) { it.receiptNo = i + 1; });
    var warnings = [];
    var zip = new JSZip(), dir = zip.folder(opt.batch);
    dir.file(opt.batch + '.xlsx', liquidationBook(XLSX, opt, sorted));
    sorted.forEach(function (it) {
      if (it.receipt && it.receipt.data) dir.file(it.receiptNo + '.' + (it.receipt.ext || 'jpg'), it.receipt.data);
      else warnings.push('Walang resibo: ' + it.expenseId);
    });

    var legs = sorted.filter(function (it) { return it.particular === TRAVEL; }).map(function (it) {
      return { date: ymd(it.date), amount: it.amount, remarks: it.desc, from: it.from || '', to: it.to || '',
        category: OBGen.guessCategory(it.ride || it.desc), claimant: it.name };
    });
    var forms = OBGen.groupForms(legs, '').filter(function (f) {
      if (f.legs.length <= OBGen.MAX_LEGS) return true;
      warnings.push('Lampas ' + OBGen.MAX_LEGS + ' rides si ' + f.claimant + ' noong ' + OBGen.longDate(f.date) + ', hindi naisama sa OB forms.');
      return false;
    });
    var ob = forms.length
      ? OBGen.buildWorkbook(JSZip, OBGen.templateBytes(), forms, opt.people || {}).then(function (bytes) {
          dir.file('OB Travel Forms - ' + longDate(opt.approvedAt) + '.xlsx', bytes);
        })
      : Promise.resolve();
    return ob.then(function () { return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }); })
      .then(function (bytes) {
        return { zip: bytes, count: sorted.length, total: sorted.reduce(function (s, it) { return s + it.amount; }, 0), obForms: forms.length, warnings: warnings };
      });
  }

  var api = { build: build, batchName: batchName };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Liquidation = api;
})(this);
