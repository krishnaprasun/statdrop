// ================================================================
//  DATA DISPATCH — Code.gs  (v2)
//  Paste into Extensions > Apps Script > Code.gs
//
//  TWO modes:
//  1. Admin sidebar  → opens inside the Google Sheet
//  2. KAM Web App    → deploy as Web App, share URL with KAMs
// ================================================================

// ── KAM Email Directory ──────────────────────────────────────────
// Update this list whenever KAMs join / leave / change emails.
var KAM_EMAILS = {
  'Anubhav':   'anubhav.tripathi@classplus.co',
  'Vineet':    'vineet.singh@classplus.co',
  'Krishna':   'krishna.prasun@classplus.co',
  'Rana':      'rana.pratap@classplus.co',
  'Bhushan':   'bhushan@classplus.co',
  'Bhupendra': 'bhupendra@classplus.co',
  'Divya':     'divya.gupta1@classplus.co',
  'Urmila':    'urmila@classplus.co',
  'H-KAM':     '',                               // no email on file
  'Satish':    'satish.sharma@classplus.co',
  'Kartikay':  'kartikay@classplus.co',
  'Asif':      'asif.ali@classplus.co',
  'Kunalgarg': 'kunal.garg@classplus.co',
  'Hashiq':    'muhammed.hashiq@classplus.co',
  'Nikita':    'nikita.chandra@classplus.co',
  'Umair':     'umair@classplus.co',
  'Harshit':   'harshit.sinha@classplus.co'
};

// ── Auto-detect KAM from their Google login email ────────────────
// Called by KAM Portal on load — returns { kamName, email } or null
function detectKAMFromSession() {
  try {
    var userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) return null;
    for (var name in KAM_EMAILS) {
      if (KAM_EMAILS[name].toLowerCase() === userEmail.toLowerCase())
        return { kamName: name, email: userEmail };
    }
    return null;
  } catch(e) { return null; }
}

// ── Web App entry point ──────────────────────────────────────────
// GET  → serves the HTML (when opening the Apps Script URL directly)
// POST → REST API called from GitHub Pages frontend
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('KAMPortal')
    .setTitle('StatDrop')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  var result;
  try {
    var req = JSON.parse(e.postData.contents);
    switch (req.action) {
      case 'getConfig':  result = getConfig();              break;
      case 'getKAMOrgs': result = getKAMOrgs(req.kamName); break;
      case 'sendData':   result = sendData(req);            break;
      default: result = { error: 'Unknown action: ' + req.action };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Admin Sidebar ───────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Data Dispatch')
    .addItem('Admin: Send to KAM', 'openSidebar')
    .addToUi();
}
function openSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Data Dispatch — Admin')
    .setWidth(390);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── Shared: Config ───────────────────────────────────────────────
function getConfig() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var tagging = ss.getSheetByName('Current_Tagging');
  if (!tagging) return { error: 'Sheet "Current_Tagging" not found.' };

  var tagData = tagging.getDataRange().getValues();
  var kams = [], kamOrgCount = {};
  for (var i = 1; i < tagData.length; i++) {
    var kam = String(tagData[i][3] || '').trim();
    if (!kam || kam.toLowerCase() === 'kam') continue;
    if (!kamOrgCount[kam]) { kams.push(kam); kamOrgCount[kam] = 0; }
    kamOrgCount[kam]++;
  }
  kams.sort();

  // Month labels — normalize ALL headers through parseMonthCode first so
  // we always get consistent "Jan23" strings regardless of whether the sheet
  // stores them as text, Date objects, or Excel serials.
  var MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var months = [], seenCodes = {};
  var revSheet = ss.getSheetByName('Revenue');
  if (revSheet) {
    var hdrs = revSheet.getRange(1, 1, 1, revSheet.getLastColumn()).getValues()[0];
    for (var c = 1; c < hdrs.length; c++) {
      var code = parseMonthCode(hdrs[c]);
      if (!code || seenCodes[code]) continue;
      seenCodes[code] = true;
      var yy = Math.floor(code / 100);
      var mm = code % 100;
      months.push(MON[mm - 1] + (yy < 10 ? '0' + yy : String(yy))); // e.g. "Jan23"
    }
  }

  var candidates = ['Revenue', 'GMV', 'Transactions', 'Downloads'];
  var existing   = candidates.filter(function(n) { return !!ss.getSheetByName(n); });

  // Find earliest available month for each sheet
  var sheetAvailability = {};
  existing.forEach(function(name) {
    var src = ss.getSheetByName(name);
    if (!src) return;
    var hdrs = src.getRange(1, 1, 1, src.getLastColumn()).getValues()[0];
    var earliestCode = 999999;
    for (var c = 1; c < hdrs.length; c++) {
      if (!hdrs[c] || String(hdrs[c]).toLowerCase().indexOf('lifetime') !== -1) continue;
      var code = parseMonthCode(hdrs[c]);
      if (code && code < earliestCode) earliestCode = code;
    }
    if (earliestCode < 999999) {
      var yy = Math.floor(earliestCode / 100);
      var mm = earliestCode % 100;
      sheetAvailability[name] = MON[mm - 1] + "'" + (yy < 10 ? '0' + yy : String(yy));
    }
  });

  // Filter months to current month max and identify current month label
  var now = new Date();
  var currentYYMM = (now.getFullYear() % 100) * 100 + (now.getMonth() + 1);
  var currentMonth = '';
  months = months.filter(function(m) {
    var code = parseMonthCode(m);
    if (!code) return false;
    if (code <= currentYYMM) { currentMonth = m; return true; }
    return false;
  });

  return { kams: kams, kamEmails: KAM_EMAILS, months: months, sheets: existing, currentMonth: currentMonth, sheetAvailability: sheetAvailability };
}

// ── Shared: Orgs for a KAM ───────────────────────────────────────
function getKAMOrgs(kamName) {
  var data = SpreadsheetApp.getActiveSpreadsheet()
               .getSheetByName('Current_Tagging').getDataRange().getValues();
  var orgs = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][3] || '').trim().toLowerCase() === kamName.trim().toLowerCase())
      orgs.push({ id: data[i][0], name: String(data[i][1]||'').trim(), category: String(data[i][2]||'').trim() });
  }
  return orgs;
}

// ── Utility: column header → YYMM integer ───────────────────────
// NOTE: GMV/Downloads/Transactions sheets store month headers as Excel dates
// where the DAY component encodes the 2-digit year (e.g. Jan 23 2026 = Jan 2023).
// Revenue stores headers as plain text like "Jan23".
function parseMonthCode(header) {
  if (header === null || header === undefined || header === '') return null;

  if (header instanceof Date) {
    var d = header.getDate();        // day (encodes year in GMV/Downloads)
    var m = header.getMonth() + 1;   // actual month 1-12
    var y = header.getFullYear() % 100; // year last 2 digits

    // When day ≠ year-component: day IS the 2-digit year (e.g. Jan-23 → day=23, yr=26)
    // When day == year-component: it's a genuine current-year date (e.g. Jan-26 → day=26, yr=26)
    if (d >= 20 && d <= 31 && d !== y) {
      return d * 100 + m;   // use day as year
    }
    return y * 100 + m;     // use year component normally
  }

  if (typeof header === 'number' && header > 20000) {
    var dt = new Date(Math.round((header - 25569) * 86400000));
    if (!isNaN(dt.getTime())) return (dt.getUTCFullYear() % 100) * 100 + (dt.getUTCMonth() + 1);
    return null;
  }

  var s = String(header).trim().toLowerCase();
  var MAP = { jan:1,feb:2,mar:3,apr:4,may:5,june:6,jun:6,july:7,jul:7,aug:8,sept:9,sep:9,oct:10,nov:11,dec:12 };
  var mt = s.match(/^([a-z]+)(\d{2})$/);
  if (mt && MAP[mt[1]]) return parseInt(mt[2]) * 100 + MAP[mt[1]];
  return null;
}

function friendlyHeader(h) {
  var code = parseMonthCode(h);
  if (code) {
    var MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var yy = Math.floor(code / 100);
    var mm = code % 100;
    return MON[mm - 1] + (yy < 10 ? '0' + yy : String(yy)); // e.g. "Jan23"
  }
  // Non-month columns (lifetime etc.) — return as-is
  if (h instanceof Date) return Utilities.formatDate(h, 'UTC', 'MMM-yy');
  if (typeof h === 'number' && h > 20000) {
    var d = new Date(Math.round((h - 25569) * 86400000));
    return Utilities.formatDate(d, 'UTC', 'MMM-yy');
  }
  return String(h);
}

// ── Core: build filtered data arrays ────────────────────────────
function buildFilteredExcel(kamName, selectedSheets, startMonth, endMonth) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var startCode = parseMonthCode(startMonth);
  var endCode   = parseMonthCode(endMonth);
  if (!startCode || !endCode || startCode > endCode)
    throw new Error('Invalid month range: ' + startMonth + ' to ' + endMonth);

  var orgs = getKAMOrgs(kamName);
  if (!orgs.length) throw new Error('No orgs tagged to ' + kamName);

  var orgIdSet = {}, orgNameMap = {};
  orgs.forEach(function(o) {
    orgIdSet[o.id] = orgIdSet[String(o.id)] = true;
    orgNameMap[o.id] = orgNameMap[String(o.id)] = o.name;
  });

  // Filter each sheet into plain arrays — no spreadsheet creation needed
  var sheetsData = [];
  selectedSheets.forEach(function(sheetName) {
    var src = ss.getSheetByName(sheetName);
    if (!src) return;
    var raw = src.getDataRange().getValues();
    if (raw.length < 2) return;

    var headers = raw[0];
    var isRev   = sheetName.toLowerCase().indexOf('revenue') !== -1;

    var colIdx = [0];
    for (var c = 1; c < headers.length; c++) {
      if (String(headers[c] || '').toLowerCase().indexOf('lifetime') !== -1) continue;
      var code = parseMonthCode(headers[c]);
      if (code && code >= startCode && code <= endCode) colIdx.push(c);
    }

    var outHeader = ['Org ID', 'Org Name'];
    for (var k = 1; k < colIdx.length; k++)
      outHeader.push(friendlyHeader(headers[colIdx[k]]));

    var outRows = [outHeader];
    for (var r = 1; r < raw.length; r++) {
      var orgId = raw[r][0];
      if (orgId === null || orgId === undefined || orgId === '') continue;
      if (!orgIdSet[orgId] && !orgIdSet[String(orgId)]) continue;

      var row = [orgId, orgNameMap[orgId] || orgNameMap[String(orgId)] || ''];
      for (var k2 = 1; k2 < colIdx.length; k2++) {
        var val = raw[r][colIdx[k2]];
        if (typeof val !== 'number') val = (val === null || val === undefined) ? 0 : val;
        if (isRev && typeof val === 'number' && val !== 0)
          val = Math.round((val / 1.18) * 100) / 100;
        row.push(val);
      }
      outRows.push(row);
    }
    sheetsData.push({ name: sheetName, rows: outRows });
  });

  // Build xlsx directly in memory — no temp spreadsheet, no Drive API, no UrlFetch
  var base64 = buildXlsxDirect(sheetsData);
  return {
    base64:    base64,
    fileName:  kamName + '_Data_' + startMonth + '_to_' + endMonth + '.xlsx',
    orgCount:  orgs.length
  };
}

// ── Direct xlsx builder ──────────────────────────────────────────
// Constructs a valid .xlsx (OpenXML) zip in memory without any
// SpreadsheetApp calls. Saves ~5-8 seconds per send.
function buildXlsxDirect(sheetsData) {

  var xmlEsc = function(s) {
    return String(s)
      .replace(/&/g,'&amp;').replace(/</g,'&lt;')
      .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  };

  // Column letter from 0-based index (A, B, …, Z, AA, …)
  var colLetter = function(n) {
    var s = ''; n++;
    while (n > 0) { s = String.fromCharCode(64 + n % 26 || 90) + s; n = Math.floor((n - 1) / 26); }
    return s;
  };

  // Shared strings table (strings are stored once, cells reference by index)
  var strings = [], strIdx = {};
  var getSS = function(val) {
    var s = String(val);
    if (strIdx[s] === undefined) { strIdx[s] = strings.length; strings.push(s); }
    return strIdx[s];
  };

  // Build one worksheet XML
  var makeSheet = function(rows) {
    var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>';
    rows.forEach(function(row, ri) {
      var cells = '';
      row.forEach(function(val, ci) {
        if (val === null || val === undefined || val === '') return;
        var ref = colLetter(ci) + (ri + 1);
        if (typeof val === 'number') {
          cells += '<c r="' + ref + '"><v>' + val + '</v></c>';
        } else {
          cells += '<c r="' + ref + '" t="s"><v>' + getSS(val) + '</v></c>';
        }
      });
      if (cells) xml += '<row r="' + (ri + 1) + '">' + cells + '</row>';
    });
    return xml + '</sheetData></worksheet>';
  };

  var worksheetXmls = sheetsData.map(function(sd) { return makeSheet(sd.rows); });

  // Shared strings XML (built after all sheets so all strings are collected)
  var ssXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' +
    strings.length + '" uniqueCount="' + strings.length + '">' +
    strings.map(function(s){ return '<si><t xml:space="preserve">' + xmlEsc(s) + '</t></si>'; }).join('') +
    '</sst>';

  // Workbook
  var wbXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
    sheetsData.map(function(sd, i) {
      return '<sheet name="' + xmlEsc(sd.name) + '" sheetId="' + (i+1) + '" r:id="rId' + (i+1) + '"/>';
    }).join('') + '</sheets></workbook>';

  // Workbook relationships
  var wbRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    sheetsData.map(function(sd, i) {
      return '<Relationship Id="rId' + (i+1) +
        '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"' +
        ' Target="worksheets/sheet' + (i+1) + '.xml"/>';
    }).join('') +
    '<Relationship Id="rIdSS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
    '<Relationship Id="rIdSt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '</Relationships>';

  // Content types
  var ctXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    sheetsData.map(function(sd, i) {
      return '<Override PartName="/xl/worksheets/sheet' + (i+1) +
        '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
    }).join('') +
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    '</Types>';

  // Root relationships
  var relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '</Relationships>';

  // Minimal styles (required by Excel to open the file)
  var stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
    '</styleSheet>';

  // Zip everything together
  var blobs = [
    Utilities.newBlob(ctXml,     'text/xml', '[Content_Types].xml'),
    Utilities.newBlob(relsXml,   'text/xml', '_rels/.rels'),
    Utilities.newBlob(wbXml,     'text/xml', 'xl/workbook.xml'),
    Utilities.newBlob(wbRelsXml, 'text/xml', 'xl/_rels/workbook.xml.rels'),
    Utilities.newBlob(ssXml,     'text/xml', 'xl/sharedStrings.xml'),
    Utilities.newBlob(stylesXml, 'text/xml', 'xl/styles.xml')
  ];
  worksheetXmls.forEach(function(xml, i) {
    blobs.push(Utilities.newBlob(xml, 'text/xml', 'xl/worksheets/sheet' + (i+1) + '.xml'));
  });

  return Utilities.base64Encode(Utilities.zip(blobs).getBytes());
}

// ── KAM Portal: download handler ────────────────────────────────
function generateExcelForDownload(params) {
  try {
    var r = buildFilteredExcel(params.kamName, params.sheets, params.startMonth, params.endMonth);
    return { success: true, base64: r.base64, fileName: r.fileName, orgCount: r.orgCount };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── Admin: email handler ─────────────────────────────────────────
function sendData(params) {
  try {
    // Use stored email if the admin left the field blank
    var email = (params.email && params.email.trim())
      ? params.email.trim()
      : KAM_EMAILS[params.kamName] || '';
    if (!email) return { success: false, message: 'No email found for ' + params.kamName + '. Please enter it manually.' };

    var r = buildFilteredExcel(params.kamName, params.sheets, params.startMonth, params.endMonth);
    var blob = Utilities.newBlob(
      Utilities.base64Decode(r.base64),
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      r.fileName
    );
    var period = params.startMonth + ' to ' + params.endMonth;
    GmailApp.sendEmail(email,
      'Data Report — ' + period + ' | ' + params.kamName,
      'Hi ' + params.kamName + ',\n\nPlease find your data report for ' + period + ' attached.\n\n' +
      'Sheets: ' + params.sheets.join(', ') + '\nOrganisations: ' + r.orgCount + '\n\n' +
      'Note: Revenue is net (ex-GST).\n\nRegards',
      { attachments: [blob], name: 'StatDrop' }
    );
    return { success: true, message: 'Sent to ' + email + ' — ' + r.orgCount + ' org(s), ' + params.sheets.length + ' sheet(s).' };
  } catch(e) { return { success: false, message: e.message }; }
}
