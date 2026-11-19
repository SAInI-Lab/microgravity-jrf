/**
 * JRF application backend: SAInI's Lab, IISc (ISRO microgravity project).
 *
 * Bound to a Google Sheet. Each submission from the application site becomes
 * (or updates) one row, the lab gets an email with the full application, and
 * the applicant gets a confirmation.
 *
 * Setup: see README.md in this folder.
 */

const CONFIG = {
  notify: 'amruthbhat@iisc.ac.in',   // who receives each application
  cc: '',                             // optional, comma-separated
  sheetName: 'Applications',
  confirmApplicant: true,             // send the applicant a receipt
  tz: 'Asia/Kolkata',
  maxField: 6000                      // characters kept per cell
};

const HEADERS = [
  'Received (IST)', 'Reference', 'Revision', 'Name', 'Email', 'Phone', 'Current position',
  'Highest degree', 'Institution', 'Fellowship exam', 'CV link', 'Work link', 'Why that link',
  'T-15 (Q1)', 'T-14 (Q2)', 'T-13 (Q3)', 'T-12 (Q4)', 'T-11 (Q5)', 'T-10 (Q6)', 'T-9 (Q7)',
  'T-8 (Q8)', 'T-7 (Q9)', 'T-6 (Q10)', 'T-5 (Q11)', 'T-4 (Q12)', 'T-3 (Q13)', 'T-2 (Q14)', 'T-1 (Q15)',
  'AI declaration', 'AI details', 'AI confirmed', 'Minutes on form', 'Full application'
];

function doGet() {
  return json_({ ok: true, service: 'jrf-applications', time: now_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // Bots fill the hidden "website" field; humans never see it.
    if (d.website) return json_({ ok: true, id: d.id || 'ok', received: now_() });

    const a = d.applicant || {};
    const email = clean_(a.email, 200).toLowerCase();
    const name = clean_(a.name, 200);
    if (!name || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return json_({ ok: false, error: 'missing_name_or_email' });
    const id = /^JRF-[A-Z0-9-]{4,40}$/.test(d.id || '') ? d.id : 'JRF-' + Utilities.getUuid().slice(0, 8).toUpperCase();

    // One submission per reference per 20 s (double clicks, retries).
    const cache = CacheService.getScriptCache();
    if (cache.get('seen:' + id + ':' + d.revision)) return json_({ ok: true, id: id, received: now_() });
    cache.put('seen:' + id + ':' + d.revision, '1', 20);

    lock.waitLock(20000);
    const sh = sheet_();
    const answers = (d.answers || []).slice(0, 15).map(x => clean_(x && x.answer));
    while (answers.length < 15) answers.push('');
    const ai = d.ai || {};
    const links = d.links || {};
    const received = now_();
    const row = [
      received, id, Number(d.revision) || 1, name, email, clean_(a.phone, 60), clean_(a.position, 200),
      clean_(a.degree, 300), clean_(a.institution, 300), clean_(a.exam, 100),
      clean_(links.cv, 1000), clean_(links.work, 1000), clean_(links.why, 300)
    ].concat(answers, [
      clean_(ai.use, 300), clean_(ai.details, 1500), ai.confirmed ? 'yes' : 'no',
      Number(d.elapsedMin) || '', clean_(d.markdown, 45000)
    ]);

    // Latest version wins: overwrite the row with the same reference.
    const last = sh.getLastRow();
    let at = 0;
    if (last > 1) {
      const ids = sh.getRange(2, 2, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) { at = i + 2; break; }
    }
    if (at) sh.getRange(at, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    lock.releaseLock();

    const rev = (Number(d.revision) || 1) > 1 ? ' (revised)' : '';
    MailApp.sendEmail({
      to: CONFIG.notify,
      cc: CONFIG.cc || undefined,
      replyTo: email,
      name: 'JRF applications',
      subject: 'JRF application' + rev + ': ' + name + ' [' + id + ']',
      body: clean_(d.markdown, 45000) + '\n\n---\nReference ' + id + ', received ' + received + ' IST.\nSheet: ' + book_().getUrl()
    });

    if (CONFIG.confirmApplicant) {
      MailApp.sendEmail({
        to: email,
        replyTo: CONFIG.notify,
        name: "SAInI's Lab, IISc",
        subject: 'We received your JRF application [' + id + ']',
        body: 'Hi ' + name.split(' ')[0] + ',\n\n' +
          'Your application for the Junior Research Fellowship (aging, microgravity and drug screening; ISRO collaboration) reached us on ' + received + ' IST.\n\n' +
          'Reference: ' + id + '\n\n' +
          'If you edit your answers and launch again, we keep only the latest version. We read every application ourselves.\n\n' +
          "SAInI's Lab, Indian Institute of Science, Bengaluru\n\n" +
          '---\nA copy of what you sent:\n\n' + clean_(d.markdown, 45000)
      });
    }

    return json_({ ok: true, id: id, received: received });
  } catch (err) {
    try { lock.releaseLock(); } catch (x) {}
    console.error(err);
    return json_({ ok: false, error: 'server_error' });
  }
}

/* ---------- helpers ---------- */
function book_() {
  // Works whether the script is attached to a Sheet or standalone.
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const created = SpreadsheetApp.create('JRF applications (microgravity)');
  props.setProperty('SHEET_ID', created.getId());
  return created;
}
function sheet_() {
  const ss = book_();
  let sh = ss.getSheetByName(CONFIG.sheetName);
  if (!sh) sh = ss.insertSheet(CONFIG.sheetName);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.setFrozenColumns(4);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setWrap(true);
    sh.setColumnWidths(14, 15, 320);
  }
  return sh;
}
function clean_(v, max) {
  let s = String(v == null ? '' : v).replace(/\u0000/g, '').trim().slice(0, max || CONFIG.maxField);
  if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;   // no formula injection into the sheet
  return s;
}
function now_() { return Utilities.formatDate(new Date(), CONFIG.tz, 'yyyy-MM-dd HH:mm'); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** Run once from the editor to grant permissions and create the sheet. */
function setup() {
  sheet_();
  MailApp.getRemainingDailyQuota();
  console.log('Ready. Sheet: ' + book_().getUrl());
}
