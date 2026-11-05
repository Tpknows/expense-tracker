/* OB Travel Form core (EPCF-2021-02 Request + EPCF-2021-03 Expense Report).
   Kinopya mula sa OB Travel Form Generator (https://claude.ai/artifact/Wo7YBBdtTHxdZZPreTZ4dX) noong Oct 30, 2026.
   Kapag binago ang generator, kopyahin ulit ito para pareho ang lalabas na form. */
/* OB Travel Form generator core. Works in the browser (globals XLSX, JSZip) and in Node (tests). */
(function (root) {
  'use strict';

  var TAG_RE = /^\s*travel\s*\/\s*transport(ation)?\s*$/i;
  var MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var MONTHS_LONG = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

  // Checkbox rectangles on the Travel Request tab: [col, colOff, row(0-based), rowOff]
  var BOX = {
    grab: [0, 205740, 19, 30480],    // Grab/Taxi
    public: [0, 205740, 20, 30480],  // Inland Public Transport
    private: [2, 205740, 19, 30480]  // Private Vehicle
  };

  var CATEGORIES = {
    grab: 'Grab/Taxi',
    moto: 'Moto-Taxi',
    public: 'Public Transport',
    gas: 'Gasoline',
    toll: 'Toll/Parking'
  };

  // ---------- small helpers ----------
  function clean(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); }
  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function titleCase(s) {
    return clean(s).toLowerCase().replace(/(^|[\s\-'.])([a-zñ])/g, function (m, a, b) { return a + b.toUpperCase(); });
  }
  function nameKey(s) { return clean(s).toUpperCase(); }
  function ymdKey(d) { return d.y + '-' + String(d.m).padStart(2, '0') + '-' + String(d.d).padStart(2, '0'); }
  function serial(d) { return (Date.UTC(d.y, d.m - 1, d.d) - Date.UTC(1899, 11, 30)) / 86400000; }
  function shortDate(d) { return MONTHS[d.m - 1] + ' ' + d.d; }
  function longDate(d) { return MONTHS_LONG[d.m - 1] + ' ' + d.d + ', ' + d.y; }
  function initials(name) {
    return clean(name).split(' ').filter(Boolean).map(function (w) { return w[0]; }).join('').toUpperCase().slice(0, 4);
  }

  function toYMD(v, XLSX) {
    if (v == null || v === '') return null;
    if (typeof v === 'number' && v > 20000 && v < 80000) {
      var p = XLSX.SSF.parse_date_code(v);
      return p ? { y: p.y, m: p.m, d: p.d } : null;
    }
    if (v instanceof Date && !isNaN(v)) return { y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate() };
    var s = clean(v), m;
    if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return { y: +m[1], m: +m[2], d: +m[3] };
    if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/))) {
      var y = +m[3]; if (y < 100) y += 2000;
      return { y: y, m: +m[1], d: +m[2] };
    }
    return null;
  }

  function toAmount(v) {
    if (typeof v === 'number') return v;
    var s = clean(v).replace(/[₱,\s]/g, '').replace(/^php/i, '');
    return /^-?\d+(\.\d+)?$/.test(s) ? parseFloat(s) : null;
  }

  // ---------- route + category reading ----------
  function tidyPlace(s) {
    return clean(s)
      .replace(/\s+-\s+.*$/, '')        // "Asus HQ - Allan" -> "Asus HQ"
      .replace(/\s*\([^)]*\)\s*$/, '')  // trailing "(names)"
      .replace(/^(the)\s+/i, '')
      .replace(/[.,;:]+$/, '')
      .trim();
  }

  function parseRoute(text) {
    var s = clean(text)
      .replace(/^(transpo(rtation)?|grab|taxi|angkas|joy ?ride|indrive|move ?it|lalamove)\b\s*/i, '')
      .replace(/^[-:–]\s*/, '');
    var m;
    if ((m = s.match(/\bfrom\s+(.+?)\s+(?:going\s+(?:back\s+)?)?(?:to|papunta(?:\s+sa)?)\s+(.+)$/i)))
      return { from: tidyPlace(m[1]), to: tidyPlace(m[2]) };
    if ((m = s.match(/\b(?:going|papunta|balik)\s+(?:back\s+)?(?:to|sa)\s+(.+?)\s+from\s+(.+)$/i)))
      return { from: tidyPlace(m[2]), to: tidyPlace(m[1]) };
    if ((m = s.match(/\b(?:going|papunta|balik)\s+(?:back\s+)?(?:to|sa)\s+(.+)$/i)))
      return { from: '', to: tidyPlace(m[1]) };
    if ((m = s.match(/^to\s+(.+)$/i)))
      return { from: '', to: tidyPlace(m[1]) };
    if ((m = s.match(/^([A-Za-z0-9 .'&]+?)\s+to\s+([A-Za-z0-9 .'&\-()]+)$/)) && m[1].split(' ').length <= 4)
      return { from: tidyPlace(m[1]), to: tidyPlace(m[2]) };
    return { from: '', to: '' };
  }

  // What is left of a remark once the ride words and the route are taken out.
  function purposeOf(text) {
    var s = clean(text)
      .replace(/^(transpo(rtation)?|grab|taxi|angkas|joy ?ride|indrive|move ?it)\b\s*/i, '')
      .replace(/\([^)]*\)/g, ' ')
      .replace(/\bjoy ?ride\b|\bindrive\b|\bgrab\b/ig, ' ')
      .replace(/\bfrom\s+.+?\s+(?:going\s+(?:back\s+)?)?(?:to|papunta(?:\s+sa)?)\s+[^-;]+/i, ' ')
      .replace(/\b(?:going|papunta|balik)\s+(?:back\s+)?(?:to|sa)\s+[^-;]+/i, ' ')
      .replace(/^[-:–\s]*to\s+[^-;]+/i, ' ')
      .replace(/\s+-\s+[A-Za-z]+\s*$/, '')
      .replace(/^[-:–\s]+|[-:–\s]+$/g, '');
    s = clean(s);
    if (/^[A-Za-z0-9 .'&]+\s+to\s+[A-Za-z0-9 .'&]+$/.test(s) && s.split(' ').length <= 6) s = '';
    return /[A-Za-z]{3,}/.test(s) ? s : '';
  }

  function guessCategory(text) {
    var s = clean(text).toLowerCase();
    if (/\btoll\b|parking|autosweep|easytrip|rfid/.test(s)) return 'toll';
    if (/\bgas\b|gasoline|\bfuel\b|diesel/.test(s)) return 'gas';
    if (/angkas|joy ?ride|move ?it|moto(r)?[- ]?taxi|habal/.test(s)) return 'moto';
    if (/jeep|\bbus\b|\bmrt\b|\blrt\b|\btrain\b|tricycle|trike|ferry|\bfx\b|uv express|pedicab/.test(s)) return 'public';
    return 'grab';
  }

  // ---------- reading the liquidation workbook ----------
  function readLiquidation(XLSX, data) {
    var wb = XLSX.read(data, { type: 'array', cellDates: false });
    var lines = [], notes = [], seenIds = {};
    wb.SheetNames.forEach(function (sn) {
      var aoa = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: null });
      var hdr = -1;
      for (var i = 0; i < Math.min(aoa.length, 30); i++) {
        if ((aoa[i] || []).some(function (v) { return clean(v).toUpperCase() === 'AMOUNT'; })) { hdr = i; break; }
      }
      if (hdr < 0) return;
      var H = aoa[hdr].map(function (v) { return clean(v).toUpperCase(); });
      var amountCol = H.indexOf('AMOUNT');
      var dateCol = H.findIndex(function (h) { return h === 'DATE' || h === 'DATE OF EXPENSE'; });
      var receiptCol = H.findIndex(function (h) { return h.indexOf('RECEIPT') === 0; });
      var nameCol = H.findIndex(function (h) { return /^(NAME|CLAIMANT|EMPLOYEE|REQUESTOR|PAID BY|EXPENSE BY|CREATIVE)/.test(h); });
      // Expense Log app files carry From/To/Ride typed by the claimant; ExpenseID stops one ride counting twice across tabs
      var fromCol = H.indexOf('FROM'), toCol = H.indexOf('TO'), rideCol = H.indexOf('RIDE'), idCol = H.indexOf('EXPENSEID');

      // the tag column is the one that holds "Travel/Transportation"
      var tagCounts = {};
      aoa.slice(hdr + 1).forEach(function (r) {
        (r || []).forEach(function (v, c) { if (typeof v === 'string' && TAG_RE.test(v)) tagCounts[c] = (tagCounts[c] || 0) + 1; });
      });
      var tagCol = Object.keys(tagCounts).sort(function (a, b) { return tagCounts[b] - tagCounts[a]; })[0];
      if (tagCol == null) return;
      tagCol = +tagCol;
      var remarksCol = H.findIndex(function (h, c) { return c !== tagCol && /^(REMARKS|PARTICULARS?|DETAILS|DESCRIPTION)/.test(h); });
      var used = [amountCol, dateCol, receiptCol, tagCol, remarksCol].filter(function (c) { return c >= 0; });
      var lastUsed = Math.max.apply(null, used);

      for (var r = hdr + 1; r < aoa.length; r++) {
        var row = aoa[r] || [];
        if (!(typeof row[tagCol] === 'string' && TAG_RE.test(row[tagCol]))) continue;
        if (idCol >= 0 && clean(row[idCol])) {
          if (seenIds[clean(row[idCol])]) continue;
          seenIds[clean(row[idCol])] = 1;
        }
        var amount = toAmount(row[amountCol]);
        var date = dateCol >= 0 ? toYMD(row[dateCol], XLSX) : null;
        if (!date) date = row.map(function (v) { return toYMD(v, XLSX); }).filter(Boolean)[0] || null;
        var claimant = '';
        if (nameCol >= 0) claimant = clean(row[nameCol]);
        else for (var c = row.length - 1; c > lastUsed; c--) {
          if (typeof row[c] === 'string' && clean(row[c]) && !TAG_RE.test(row[c])) { claimant = clean(row[c]); break; }
        }
        var line = {
          sheet: sn, row: r + 1, amount: amount, date: date,
          remarks: clean(remarksCol >= 0 ? row[remarksCol] : ''),
          receipt: receiptCol >= 0 && row[receiptCol] != null ? clean(row[receiptCol]) : '',
          claimant: claimant
        };
        if (amount == null || !date) {
          notes.push({ sheet: sn, row: r + 1, remarks: line.remarks, problem: amount == null ? 'no amount' : 'no readable date' });
          continue;
        }
        var route = parseRoute(line.remarks);
        line.from = fromCol >= 0 && clean(row[fromCol]) ? clean(row[fromCol]) : route.from;
        line.to = toCol >= 0 && clean(row[toCol]) ? clean(row[toCol]) : route.to;
        line.category = guessCategory(rideCol >= 0 && clean(row[rideCol]) ? row[rideCol] : line.remarks);
        lines.push(line);
      }
    });
    return { lines: lines, skipped: notes };
  }

  // One form per date per claimant. Lines without a claimant use fallbackName.
  function groupForms(lines, fallbackName) {
    var map = {};
    lines.forEach(function (l) {
      var who = l.claimant || fallbackName || '';
      var key = ymdKey(l.date) + '|' + nameKey(who);
      if (!map[key]) map[key] = { key: key, date: l.date, claimantKey: nameKey(who), claimant: who ? titleCase(who) : '', legs: [] };
      map[key].legs.push(l);
    });
    return Object.keys(map).sort().map(function (k) {
      var f = map[k];
      f.total = f.legs.reduce(function (s, l) { return s + l.amount; }, 0);
      var tos = [];
      f.legs.forEach(function (l) { if (l.to && tos.indexOf(l.to) < 0) tos.push(l.to); });
      f.destination = tos.join(' / ');
      var purposes = [];
      f.legs.forEach(function (l) {
        var p = purposeOf(l.remarks);
        if (p && purposes.indexOf(p) < 0) purposes.push(p);
      });
      f.purpose = purposes.join('; ');
      return f;
    });
  }

  // ---------- XML cell editing ----------
  function colNum(col) { var n = 0; for (var i = 0; i < col.length; i++) n = n * 26 + col.charCodeAt(i) - 64; return n; }

  function buildCell(ref, sAttr, val) {
    if (val == null || val === '') return '<c r="' + ref + '"' + sAttr + '/>';
    if (typeof val === 'object' && val.f) return '<c r="' + ref + '"' + sAttr + '><f>' + esc(val.f) + '</f></c>';
    if (typeof val === 'number') return '<c r="' + ref + '"' + sAttr + '><v>' + val + '</v></c>';
    return '<c r="' + ref + '"' + sAttr + ' t="inlineStr"><is><t xml:space="preserve">' + esc(val) + '</t></is></c>';
  }

  function setCell(xml, ref, val) {
    var col = ref.match(/[A-Z]+/)[0], rowN = +ref.match(/\d+/)[0];
    var cellRe = new RegExp('<c r="' + ref + '"(\\s[^>]*?)?(?:/>|>[\\s\\S]*?</c>)');
    var m = xml.match(cellRe);
    if (m) {
      var s = ((m[1] || '').match(/\ss="\d+"/) || [''])[0];
      return xml.replace(cellRe, function () { return buildCell(ref, s, val); });
    }
    var cell = buildCell(ref, '', val);
    var rowRe = new RegExp('<row r="' + rowN + '"([^>]*?)(/>|>([\\s\\S]*?)</row>)');
    var rm = xml.match(rowRe);
    if (rm) {
      var inner = rm[3] || '', target = colNum(col), cells = inner.match(/<c r="[A-Z]+\d+"[\s\S]*?(?:\/>|<\/c>)/g) || [];
      var at = cells.findIndex(function (c) { return colNum(c.match(/<c r="([A-Z]+)/)[1]) > target; });
      if (at < 0) cells.push(cell); else cells.splice(at, 0, cell);
      return xml.replace(rowRe, function () { return '<row r="' + rowN + '"' + rm[1] + '>' + cells.join('') + '</row>'; });
    }
    var rows = xml.match(/<row r="(\d+)"/g) || [], after = null;
    rows.forEach(function (r) { var n = +r.match(/\d+/)[0]; if (n > rowN && after == null) after = n; });
    var newRow = '<row r="' + rowN + '">' + cell + '</row>';
    if (after != null) return xml.replace('<row r="' + after + '"', newRow + '<row r="' + after + '"');
    return xml.replace('</sheetData>', newRow + '</sheetData>').replace('<sheetData/>', '<sheetData>' + newRow + '</sheetData>');
  }

  function checkmark(box, id) {
    var col = box[0], colOff = Math.max(0, box[1] - 15000), row = box[2], rowOff = 0;
    return '<xdr:oneCellAnchor><xdr:from><xdr:col>' + col + '</xdr:col><xdr:colOff>' + colOff + '</xdr:colOff><xdr:row>' + row +
      '</xdr:row><xdr:rowOff>' + rowOff + '</xdr:rowOff></xdr:from><xdr:ext cx="230000" cy="230000"/><xdr:pic><xdr:nvPicPr>' +
      '<xdr:cNvPr id="' + id + '" name="Check ' + id + '" descr="Checkmark"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>' +
      '<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId4"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
      '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="230000" cy="230000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>';
  }

  // ---------- filling one form ----------
  var ITIN_ROWS = [11, 12, 13, 14, 15, 16, 17], ITIN_ROWS2 = [55, 56, 57, 58, 59, 60, 61];
  var EXP_ROWS = [21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32], EXP_ROWS2 = [65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76];
  var MAX_LEGS = 14;

  function fillRequest(xml, f, person) {
    var d = serial(f.date);
    var sums = { grab: 0, moto: 0, public: 0, gas: 0, toll: 0 };
    f.legs.forEach(function (l) { sums[l.category] += l.amount; });
    // Moto-Taxi has no row of its own on the form; it is claimed under Grab/Taxi Fee
    sums.grab += sums.moto;
    xml = setCell(xml, 'B4', person.name);
    xml = setCell(xml, 'B5', person.id ? (/^\d+$/.test(person.id) ? +person.id : person.id) : '');
    xml = setCell(xml, 'E4', d);
    xml = setCell(xml, 'E5', person.department || '');
    xml = setCell(xml, 'B8', d);
    xml = setCell(xml, 'E8', f.destination || '');
    xml = setCell(xml, 'B9', f.purpose || '');
    if (sums.grab) xml = setCell(xml, 'F34', sums.grab);
    if (sums.public) xml = setCell(xml, 'F35', sums.public);
    if (sums.gas) xml = setCell(xml, 'F38', sums.gas);
    if (sums.toll) { xml = setCell(xml, 'A43', 'Toll / Parking fees'); xml = setCell(xml, 'F43', sums.toll); }
    xml = setCell(xml, 'B53', person.name);
    xml = setCell(xml, 'B55', person.superior || '');
    xml = setCell(xml, 'B56', person.head || '');
    return { xml: xml, sums: sums };
  }

  function fillReport(xml, f, person) {
    var d = serial(f.date);
    xml = setCell(xml, 'B4', person.name);
    xml = setCell(xml, 'B5', person.superior || '');
    xml = setCell(xml, 'E5', person.department || '');
    xml = setCell(xml, 'B6', d);
    xml = setCell(xml, 'E6', f.purpose || '');
    xml = setCell(xml, 'C20', 'Category');
    var spill = f.legs.length > ITIN_ROWS.length;
    f.legs.forEach(function (l, i) {
      var ir = i < 7 ? ITIN_ROWS[i] : ITIN_ROWS2[i - 7];
      xml = setCell(xml, 'B' + ir, d);
      xml = setCell(xml, 'D' + ir, l.from || '');
      xml = setCell(xml, 'E' + ir, l.to || '');
      // keep the expense lines on the same page as their itinerary rows
      var er = i < 7 ? EXP_ROWS[i] : EXP_ROWS2[i - 7];
      xml = setCell(xml, 'B' + er, d);
      xml = setCell(xml, 'C' + er, CATEGORIES[l.category]);
      xml = setCell(xml, 'D' + er, l.remarks);
      xml = setCell(xml, 'E' + er, l.amount);
    });
    if (spill) {
      xml = setCell(xml, 'B48', person.name);
      xml = setCell(xml, 'B49', person.superior || '');
      xml = setCell(xml, 'E49', person.department || '');
      xml = setCell(xml, 'B50', d);
      xml = setCell(xml, 'E50', f.purpose || '');
      xml = setCell(xml, 'C64', 'Category');
      xml = setCell(xml, 'E77', { f: 'SUM(E65:E76)' });
      xml = setCell(xml, 'E78', { f: 'E35' });
      xml = setCell(xml, 'E80', { f: 'E77+E78-E79' });
    }
    return xml;
  }

  function sheetNames(forms) {
    var seen = {};
    return forms.map(function (f) {
      var base = shortDate(f.date) + (f.claimant ? ' ' + initials(f.claimant) : '');
      var n = base, k = 2;
      while (seen[n]) n = base + ' (' + (k++) + ')';
      seen[n] = 1;
      return { req: (n + ' Request').slice(0, 31), rep: (n + ' Report').slice(0, 31) };
    });
  }

  // ---------- building the workbook ----------
  function buildWorkbook(JSZip, templateBytes, forms, people) {
    return JSZip.loadAsync(templateBytes).then(function (zip) {
      var names = ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/drawings/drawing1.xml', 'xl/drawings/drawing2.xml',
        'xl/drawings/_rels/drawing1.xml.rels', 'xl/drawings/_rels/drawing2.xml.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml'];
      return Promise.all(names.map(function (n) { return zip.file(n).async('string'); })).then(function (t) {
        return Promise.all([zip.file('xl/printerSettings/printerSettings1.bin').async('uint8array'),
          zip.file('xl/printerSettings/printerSettings2.bin').async('uint8array')]).then(function (ps) {
          var T = {}; names.forEach(function (n, i) { T[n] = t[i]; });
          var labels = sheetNames(forms);
          var sheetsXml = '', relsXml = '', ctXml = '';
          var wsRels = function (k) {
            return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
              '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing' + k + '.xml"/>' +
              '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="../printerSettings/printerSettings' + k + '.bin"/></Relationships>';
          };
          var checkId = 300;
          forms.forEach(function (f, i) {
            var person = people[f.claimantKey] || { name: f.claimant };
            person = Object.assign({}, person, { name: person.name || f.claimant });
            var a = 2 * i + 1, b = 2 * i + 2;
            var req = fillRequest(T['xl/worksheets/sheet1.xml'], f, person);
            var rep = fillReport(T['xl/worksheets/sheet2.xml'], f, person);
            if (i === 0) req.xml = req.xml.replace('<sheetView ', '<sheetView tabSelected="1" ');
            var marks = '';
            if (req.sums.grab) marks += checkmark(BOX.grab, checkId++);
            if (req.sums.public) marks += checkmark(BOX.public, checkId++);
            if (req.sums.gas) marks += checkmark(BOX.private, checkId++);
            zip.file('xl/worksheets/sheet' + a + '.xml', req.xml);
            zip.file('xl/worksheets/sheet' + b + '.xml', rep);
            zip.file('xl/worksheets/_rels/sheet' + a + '.xml.rels', wsRels(a));
            zip.file('xl/worksheets/_rels/sheet' + b + '.xml.rels', wsRels(b));
            zip.file('xl/drawings/drawing' + a + '.xml', T['xl/drawings/drawing1.xml'].replace('</xdr:wsDr>', marks + '</xdr:wsDr>'));
            zip.file('xl/drawings/drawing' + b + '.xml', T['xl/drawings/drawing2.xml']);
            zip.file('xl/drawings/_rels/drawing' + a + '.xml.rels', T['xl/drawings/_rels/drawing1.xml.rels']);
            zip.file('xl/drawings/_rels/drawing' + b + '.xml.rels', T['xl/drawings/_rels/drawing2.xml.rels']);
            zip.file('xl/printerSettings/printerSettings' + a + '.bin', ps[0]);
            zip.file('xl/printerSettings/printerSettings' + b + '.bin', ps[1]);
            sheetsXml += '<sheet name="' + esc(labels[i].req) + '" sheetId="' + a + '" r:id="rIdS' + a + '"/>' +
              '<sheet name="' + esc(labels[i].rep) + '" sheetId="' + b + '" r:id="rIdS' + b + '"/>';
            [a, b].forEach(function (k) {
              relsXml += '<Relationship Id="rIdS' + k + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + k + '.xml"/>';
              ctXml += '<Override PartName="/xl/worksheets/sheet' + k + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
                '<Override PartName="/xl/drawings/drawing' + k + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>';
            });
          });
          zip.file('xl/workbook.xml', T['xl/workbook.xml'].replace(/<sheets>[\s\S]*?<\/sheets>/, '<sheets>' + sheetsXml + '</sheets>'));
          zip.file('xl/_rels/workbook.xml.rels', T['xl/_rels/workbook.xml.rels']
            .replace(/<Relationship [^>]*relationships\/worksheet"[^>]*\/>/g, '')
            .replace('</Relationships>', relsXml + '</Relationships>'));
          zip.file('[Content_Types].xml', T['[Content_Types].xml']
            .replace(/<Override PartName="\/xl\/(worksheets|drawings)\/[^"]+"[^>]*\/>/g, '')
            .replace('</Types>', ctXml + '</Types>'));
          return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
        });
      });
    });
  }

  var api = {
    readLiquidation: readLiquidation, groupForms: groupForms, buildWorkbook: buildWorkbook,
    parseRoute: parseRoute, guessCategory: guessCategory, titleCase: titleCase, nameKey: nameKey,
    shortDate: shortDate, longDate: longDate, CATEGORIES: CATEGORIES, MAX_LEGS: MAX_LEGS
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.OBGen = api;
})(this);

/* EasyPC OB Travel template (xlsx), base64 */
OBGen.TEMPLATE_B64='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';
OBGen.templateBytes=function(){var s=atob(OBGen.TEMPLATE_B64),a=new Uint8Array(s.length);for(var i=0;i<s.length;i++)a[i]=s.charCodeAt(i);return a;};
