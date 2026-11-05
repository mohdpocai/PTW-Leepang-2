/**
 * PTW – Leepang 2 Estate
 * Google Apps Script backend: stores Permit to Work records in a Google Sheet.
 *
 * SETUP
 * 1. Create a new Google Sheet (e.g. "PTW Leepang 2 Estate").
 * 2. Extensions > Apps Script. Delete the sample code, paste this whole file, Save.
 * 3. Change APPROVER_PIN below.
 * 4. Deploy > New deployment > type "Web app"
 *      Execute as: Me
 *      Who has access: Anyone
 *    Click Deploy, allow permissions, copy the Web app URL (ends with /exec).
 * 5. Paste that URL into SCRIPT_URL at the top of index.html.
 *
 * PDF: each permit is saved as a PDF in the Drive folder PDF_FOLDER and linked in the
 * "PDF" column. The first time, run the function  authorize  once (select it at the top,
 * click Run, allow Drive access), then deploy a New version.
 *
 * After editing this script later: Deploy > Manage deployments > Edit (pencil)
 * > Version: New version > Deploy  (the URL stays the same).
 */

const SHEET_NAME = 'PTW';
const APPROVER_PIN = '1234';           // <-- CHANGE THIS. Needed to Approve / Reject / Close.
const ESTATE_CODE = 'LP2';
const TZ = 'Asia/Kuala_Lumpur';
const PDF_FOLDER = 'PTW Leepang 2 - PDF';
const PDF_PUBLIC_LINK = true;          // true = anyone with the link can open the PDF (easy on phones)

const HEADERS = [
  'Permit No', 'Submitted At', 'Issue Date', 'Work Types', 'Division / Block', 'Exact Location',
  'Job Description', 'Valid From', 'Valid To', 'Applicant', 'Position', 'Company / Contractor',
  'Phone', 'No. of Workers', 'Worker Names', 'Equipment / Tools', 'Hazards', 'PPE Required',
  'Precautions', 'Gas O2 (%)', 'Gas LEL (%)', 'Gas H2S (ppm)', 'Gas CO (ppm)', 'Gas Tested By',
  'Additional Precautions', 'Emergency Contact', 'Status',
  'Approved By', 'Approved At', 'Approval Remarks',
  'Closed By', 'Closed At', 'Closure Checks', 'Closure Remarks', 'Last Updated', 'PDF'
];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold').setBackground('#1f5130').setFontColor('#ffffff').setWrap(true);
    sh.getRange(1, 1, sh.getMaxRows(), HEADERS.length).setNumberFormat('@');   // keep everything as plain text
    sh.setColumnWidth(1, 150);
  }
  // add any new columns (e.g. PDF) to an older sheet
  const lastCol = sh.getLastColumn();
  if (lastCol < HEADERS.length) {
    sh.getRange(1, lastCol + 1, 1, HEADERS.length - lastCol).setValues([HEADERS.slice(lastCol)])
      .setFontWeight('bold').setBackground('#1f5130').setFontColor('#ffffff');
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function now_() {
  return Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
}

/** GET: with ?payload= -> create / update a permit; otherwise return all permits, newest first */
function doGet(e) {
  if (e && e.parameter && e.parameter.payload) {
    return handle_(e.parameter.payload);
  }
  const sh = getSheet_();
  const values = sh.getDataRange().getDisplayValues();
  const head = values.shift() || [];
  const rows = values.filter(r => r[0]).map(r => {
    const o = {};
    head.forEach((h, i) => o[h] = r[i]);
    return o;
  });
  return json_({ ok: true, rows: rows.reverse() });
}

/** POST also supported */
function doPost(e) {
  return handle_(e.postData.contents);
}

function handle_(raw) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse(raw);
    if (body.action === 'create') return json_(create_(body.data || {}));
    if (body.action === 'update') return json_(update_(body));
    return json_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function create_(d) {
  const sh = getSheet_();
  const year = Utilities.formatDate(new Date(), TZ, 'yyyy');
  // running number per year
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat() : [];
  const prefix = 'PTW-' + ESTATE_CODE + '-' + year + '-';
  const n = ids.filter(id => String(id).indexOf(prefix) === 0).length + 1;
  const permitNo = prefix + String(n).padStart(4, '0');

  d['Permit No'] = permitNo;
  d['Submitted At'] = now_();
  d['Status'] = 'Pending';
  d['Last Updated'] = d['Submitted At'];

  sh.appendRow(HEADERS.map(h => d[h] === undefined ? '' : String(d[h])));
  const pdf = savePdf_(sh, sh.getLastRow());
  return { ok: true, permitNo: permitNo, pdf: pdf };
}

function update_(b) {
  if (String(b.pin) !== APPROVER_PIN) return { ok: false, error: 'Wrong PIN' };
  if (!b.name) return { ok: false, error: 'Name is required' };

  const sh = getSheet_();
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().flat();
  const rowIdx = ids.indexOf(b.permitNo);
  if (rowIdx < 1) return { ok: false, error: 'Permit not found' };
  const row = rowIdx + 1;
  const col = h => HEADERS.indexOf(h) + 1;
  const current = sh.getRange(row, col('Status')).getValue();
  const t = now_();

  if (b.status === 'Approved' || b.status === 'Rejected') {
    if (current !== 'Pending') return { ok: false, error: 'Permit is already ' + current };
    sh.getRange(row, col('Approved By')).setValue(b.name);
    sh.getRange(row, col('Approved At')).setValue(t);
    sh.getRange(row, col('Approval Remarks')).setValue(b.remarks || '');
  } else if (b.status === 'Closed') {
    if (current !== 'Approved') return { ok: false, error: 'Only an Approved permit can be closed' };
    sh.getRange(row, col('Closed By')).setValue(b.name);
    sh.getRange(row, col('Closed At')).setValue(t);
    sh.getRange(row, col('Closure Checks')).setValue(b.checks || '');
    sh.getRange(row, col('Closure Remarks')).setValue(b.remarks || '');
  } else {
    return { ok: false, error: 'Unknown status' };
  }
  sh.getRange(row, col('Status')).setValue(b.status);
  sh.getRange(row, col('Last Updated')).setValue(t);
  SpreadsheetApp.flush();
  const pdf = savePdf_(sh, row);
  return { ok: true, pdf: pdf };
}

/* ================= PDF ================= */

/** Run this once from the editor to give the script Drive permission. */
function authorize() {
  getFolder_();
  getSheet_();
  Logger.log('OK - Drive and Sheet access granted');
}

/** Create a PDF for every row that has no PDF yet (e.g. permits saved before this update). */
function makeMissingPdfs() {
  const sh = getSheet_();
  const pdfCol = HEADERS.indexOf('PDF') + 1;
  for (let r = 2; r <= sh.getLastRow(); r++) {
    if (sh.getRange(r, 1).getValue() && !sh.getRange(r, pdfCol).getValue()) savePdf_(sh, r);
  }
}

function getFolder_() {
  const it = DriveApp.getFoldersByName(PDF_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PDF_FOLDER);
}

/** Build the PDF for one sheet row, replace any older copy, write the link into the PDF column. */
function savePdf_(sh, row) {
  try {
    const vals = sh.getRange(row, 1, 1, HEADERS.length).getDisplayValues()[0];
    const r = {};
    HEADERS.forEach((h, i) => r[h] = vals[i]);
    const name = r['Permit No'] + '.pdf';
    const folder = getFolder_();
    const old = folder.getFilesByName(name);
    while (old.hasNext()) old.next().setTrashed(true);
    const blob = Utilities.newBlob(pdfHtml_(r), 'text/html', r['Permit No'] + '.html').getAs('application/pdf').setName(name);
    const file = folder.createFile(blob);
    if (PDF_PUBLIC_LINK) file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = file.getUrl();
    sh.getRange(row, HEADERS.indexOf('PDF') + 1).setValue(url);
    return url;
  } catch (err) {
    Logger.log('PDF failed: ' + err);
    return '';
  }
}

function pdfEsc_(v) {
  return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
}

/** Ticked / unticked list laid out in a 3-column table (PDF engine does not support grid/flex). */
function pdfChecks_(value, all) {
  const picked = String(value || '').split('; ');
  let html = '<table class="ck"><tr>';
  all.forEach((x, i) => {
    if (i && i % 3 === 0) html += '</tr><tr>';
    html += '<td>' + (picked.indexOf(x) > -1 ? '&#9745;' : '&#9744;') + ' ' + pdfEsc_(x) + '</td>';
  });
  return html + '</tr></table>';
}

const PDF_LISTS = {
  work: ['Hot Work (welding / cutting / grinding)','Working at Height','Confined Space Entry','Electrical Work',
         'Excavation / Trenching','Chemical / Pesticide Handling','Heavy Machinery / Vehicle',
         'Palm / Tree Felling','Roofing / Building Repair','General Maintenance'],
  haz:  ['Fall from height','Fire / explosion','Electric shock','Chemical / toxic exposure','Oxygen deficiency / toxic gas',
         'Moving machinery / vehicles','Falling objects (fronds, FFB, tools)','Overhead power lines','Underground cables / pipes',
         'Snake / insect / animal bites','Heat stress','Noise','Sharp tools / cuts','Slippery / uneven ground'],
  ppe:  ['Safety helmet','Safety boots','Gloves','Safety goggles / face shield','Ear protection','Respirator / mask',
         'Full body harness','Reflective vest','Chemical apron / suit','Welding mask & apron','Long sleeves & trousers'],
  pre:  ['Toolbox talk / JSA done','Workers trained & competent','Area barricaded & signage up','Equipment & tools inspected',
         'Isolation / Lock-out Tag-out done','Fire extinguisher at site','Gas test done','Standby person / watcher assigned',
         'First aid kit at site','Ladder / scaffold checked & secured','Emergency plan & assembly point briefed','Weather checked (no lightning / heavy rain)'],
  close:['Work completed','All workers out & accounted for','Tools & materials removed','Area cleaned & safe',
         'Isolations / LOTO removed','Barricades removed']
};

function pdfHtml_(r) {
  const e = pdfEsc_;
  const kv = (k, v) => '<tr><td class="k">' + k + '</td><td>' + v + '</td></tr>';
  const status = r['Status'] || 'Pending';
  const stColor = { Pending: '#8a5a00', Approved: '#1f5130', Rejected: '#b42318', Closed: '#475467' }[status] || '#000';
  const gas = (r['Gas O2 (%)'] || r['Gas LEL (%)'] || r['Gas H2S (ppm)'] || r['Gas CO (ppm)'])
    ? '<table class="t">' + kv('Gas Test', 'O2: ' + e(r['Gas O2 (%)']) + ' % &nbsp; LEL: ' + e(r['Gas LEL (%)']) + ' % &nbsp; H2S: ' +
        e(r['Gas H2S (ppm)']) + ' ppm &nbsp; CO: ' + e(r['Gas CO (ppm)']) + ' ppm &nbsp; Tested by: ' + e(r['Gas Tested By'])) + '</table>' : '';
  return '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,sans-serif;font-size:10px;color:#000;margin:0}' +
    '.hd{width:100%;border-bottom:3px solid #1f5130;margin-bottom:6px}.hd td{vertical-align:bottom;padding:0 0 4px}' +
    'h1{font-size:17px;margin:0;color:#1f5130}.no{font-size:14px;font-weight:bold;text-align:right}' +
    'h2{font-size:10.5px;background:#1f5130;color:#fff;padding:3px 6px;margin:8px 0 3px}' +
    '.t{width:100%;border-collapse:collapse}.t td{border:1px solid #999;padding:3px 5px;vertical-align:top}' +
    '.t td.k{width:26%;background:#f0f3f0;font-weight:bold}' +
    '.ck{width:100%;border-collapse:collapse}.ck td{width:33%;padding:1px 4px;vertical-align:top;font-size:9.5px}' +
    '.sig{width:100%;border-collapse:collapse;margin-top:3px}.sig td{width:33%;border:1px solid #999;padding:5px;vertical-align:top;height:78px}' +
    '.sig b{display:block;margin-bottom:3px}.note{font-size:8.5px;margin-top:8px;color:#333}' +
    '</style></head><body>' +
    '<table class="hd"><tr><td><h1>PERMIT TO WORK</h1>Leepang 2 Estate</td>' +
    '<td class="no">' + e(r['Permit No']) + '<br><span style="font-size:11px;color:' + stColor + '">Status: ' + e(status) + '</span></td></tr></table>' +
    '<h2>1. WORK DETAILS</h2><table class="t">' +
      kv('Issue Date', e(r['Issue Date'])) + kv('Division / Block', e(r['Division / Block'])) +
      kv('Exact Location', e(r['Exact Location'])) + kv('Valid From / To', e(r['Valid From']) + ' &nbsp;to&nbsp; ' + e(r['Valid To'])) +
      kv('Job Description', e(r['Job Description'])) + kv('Equipment / Tools', e(r['Equipment / Tools'])) + '</table>' +
    '<h2>2. TYPE OF WORK</h2>' + pdfChecks_(r['Work Types'], PDF_LISTS.work) +
    '<h2>3. APPLICANT &amp; WORKERS</h2><table class="t">' +
      kv('Applicant / Position', e(r['Applicant']) + ' - ' + e(r['Position'])) +
      kv('Company / Phone', e(r['Company / Contractor']) + (r['Phone'] ? ' / ' + e(r['Phone']) : '')) +
      kv('No. of Workers', e(r['No. of Workers'])) + kv('Worker Names', e(r['Worker Names'])) +
      kv('Emergency Contact', e(r['Emergency Contact'])) + '</table>' +
    '<h2>4. HAZARDS IDENTIFIED</h2>' + pdfChecks_(r['Hazards'], PDF_LISTS.haz) +
    '<h2>5. PPE REQUIRED</h2>' + pdfChecks_(r['PPE Required'], PDF_LISTS.ppe) +
    '<h2>6. PRECAUTIONS TAKEN</h2>' + pdfChecks_(r['Precautions'], PDF_LISTS.pre) + gas +
    (r['Additional Precautions'] ? '<table class="t">' + kv('Additional Precautions', e(r['Additional Precautions'])) + '</table>' : '') +
    '<h2>7. AUTHORISATION</h2><table class="sig"><tr>' +
      '<td><b>Applicant</b>' + e(r['Applicant']) + '<br>' + e(r['Submitted At']) + '<br><br>Sign: ______________</td>' +
      '<td><b>' + (status === 'Rejected' ? 'Rejected by' : 'Approved by') + '</b>' + e(r['Approved By']) + '<br>' + e(r['Approved At']) +
        '<br>' + e(r['Approval Remarks']) + '<br>Sign: ______________</td>' +
      '<td><b>Closed by</b>' + e(r['Closed By']) + '<br>' + e(r['Closed At']) + '<br>' + e(r['Closure Remarks']) + '<br>Sign: ______________</td>' +
    '</tr></table>' +
    '<h2>8. CLOSE-OUT</h2>' + pdfChecks_(r['Closure Checks'], PDF_LISTS.close) +
    '<p class="note">This permit is valid only for the work, location and period stated. Keep a copy at the work site. ' +
    'Work must stop and the permit be re-assessed if conditions change. Generated ' + e(Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm')) + '.</p>' +
    '</body></html>';
}
