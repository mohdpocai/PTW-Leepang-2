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
 * After editing this script later: Deploy > Manage deployments > Edit (pencil)
 * > Version: New version > Deploy  (the URL stays the same).
 */

const SHEET_NAME = 'PTW';
const APPROVER_PIN = '1234';           // <-- CHANGE THIS. Needed to Approve / Reject / Close.
const ESTATE_CODE = 'LP2';
const TZ = 'Asia/Kuala_Lumpur';

const HEADERS = [
  'Permit No', 'Submitted At', 'Issue Date', 'Work Types', 'Division / Block', 'Exact Location',
  'Job Description', 'Valid From', 'Valid To', 'Applicant', 'Position', 'Company / Contractor',
  'Phone', 'No. of Workers', 'Worker Names', 'Equipment / Tools', 'Hazards', 'PPE Required',
  'Precautions', 'Gas O2 (%)', 'Gas LEL (%)', 'Gas H2S (ppm)', 'Gas CO (ppm)', 'Gas Tested By',
  'Additional Precautions', 'Emergency Contact', 'Status',
  'Approved By', 'Approved At', 'Approval Remarks',
  'Closed By', 'Closed At', 'Closure Checks', 'Closure Remarks', 'Last Updated'
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
    sh.getRange('A:AI').setNumberFormat('@');   // keep everything as plain text
    sh.setColumnWidth(1, 150);
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
  return { ok: true, permitNo: permitNo };
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
  return { ok: true };
}
