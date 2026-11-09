/**
 * OR Crisis Simulator report card service
 * Receives results from the game, logs them to a Google Sheet,
 * and emails a styled report card to the instructor and, on request, to the student.
 *
 * Setup: paste into a new Apps Script project, run setup() once, then
 * Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 */

const INSTRUCTOR_EMAIL = 'bfulton.sim.director@gmail.com';
const ACCESS_CODE = 'MH2026X4V6LPPQ776XX'; // must match the code students enter in the game
const COURSE = 'ANES 1123';
const SCENARIO = 'OR Crisis Simulator';
const SHEET_NAME = 'Results';
const HEADERS = ['Submitted', 'Verification code', 'Student name', 'Score', 'Possible', 'Points earned', 'Points lost', 'Case time', 'Decisions', 'Student copy sent to', 'Student copy sent at', 'Report JSON'];

/* ---------- Run once from the editor ---------- */
function setup() {
  const ss = getSheet_().getParent();
  MailApp.sendEmail({
    to: INSTRUCTOR_EMAIL,
    subject: SCENARIO + ': report card service is ready',
    htmlBody: '<p>The report card service is set up.</p><p>Results will be logged here: <a href="' + ss.getUrl() + '">' + ss.getName() + '</a></p>'
  });
  Logger.log('Results sheet: ' + ss.getUrl());
}

/* ---------- Web app entry points ---------- */
function doGet() {
  return json_({ ok: true, service: SCENARIO + ' report cards' });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'submit') return json_(submit_(body));
    if (body.action === 'copy') return json_(copy_(body));
    return json_({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    return json_({ ok: false, error: 'The report could not be processed: ' + err.message });
  }
}

/* ---------- Actions ---------- */
function submit_(b) {
  const given = String(b.accessCode || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (given !== ACCESS_CODE) return { ok: false, error: 'This report was not accepted because the access code is missing or incorrect.' };
  const r = clean_(b);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let code;
  try {
    const sh = getSheet_();
    code = newCode_(sh);
    r.code = code;
    r.submitted = new Date();
    sh.appendRow([r.submitted, code, r.name, r.score, r.max, r.earned, r.lost, r.caseTime, r.ledger.length, '', '', JSON.stringify(r)]);
  } finally {
    lock.releaseLock();
  }
  MailApp.sendEmail({
    to: INSTRUCTOR_EMAIL,
    subject: SCENARIO + ' report card: ' + r.name + ', ' + r.score + '/' + r.max + ' (' + code + ')',
    htmlBody: reportHtml_(r, true),
    body: r.name + ' scored ' + r.score + ' of ' + r.max + '. Verification code ' + code + '.',
    name: SCENARIO
  });
  return { ok: true, code: code };
}

function copy_(b) {
  const email = String(b.email || '').trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 120) return { ok: false, error: 'That email address does not look right. Check it and try again.' };
  const code = String(b.code || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const sh = getSheet_();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === code) {
      if (rows[i][9]) return { ok: false, error: 'A copy of this report card was already sent to ' + rows[i][9] + '.' };
      const r = JSON.parse(rows[i][11]);
      r.submitted = rows[i][0];
      MailApp.sendEmail({
        to: email,
        subject: 'Your ' + SCENARIO + ' report card (' + code + ')',
        htmlBody: reportHtml_(r, false),
        body: 'Your score: ' + r.score + ' of ' + r.max + '. Verification code ' + code + '.',
        name: SCENARIO,
        replyTo: INSTRUCTOR_EMAIL
      });
      sh.getRange(i + 1, 10, 1, 2).setValues([[email, new Date()]]);
      return { ok: true };
    }
  }
  return { ok: false, error: 'That report card could not be found.' };
}

/* ---------- Helpers ---------- */
function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID');
  let ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create(SCENARIO + ' Results');
    props.setProperty('SHEET_ID', ss.getId());
  }
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function newCode_(sh) {
  const chars = 'ABCDEFGHJKMNPQRTUVWXYZ2346789';
  const used = new Set(sh.getRange(1, 2, Math.max(1, sh.getLastRow()), 1).getValues().map(r => r[0]));
  let c;
  do { c = 'OR'; for (let i = 0; i < 5; i++) c += chars.charAt(Math.floor(Math.random() * chars.length)); } while (used.has(c));
  return c;
}

function clean_(b) {
  const num = (v, lo, hi) => Math.max(lo, Math.min(hi, Math.round(Number(v) || 0)));
  const ledger = (Array.isArray(b.ledger) ? b.ledger : []).slice(0, 80).map(l => ({
    t: String(l.t || '').slice(0, 8),
    label: String(l.label || '').slice(0, 300),
    pts: num(l.pts, -100, 100)
  }));
  return {
    name: String(b.name || 'Student').trim().slice(0, 60) || 'Student',
    score: num(b.score, -1000, 1000),
    max: num(b.max, 0, 1000),
    earned: num(b.earned, 0, 1000),
    lost: num(b.lost, -1000, 0),
    caseTime: String(b.caseTime || '').slice(0, 8),
    ledger: ledger
  };
}

function esc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Email report card (email safe tables and inline styles) ---------- */
function reportHtml_(r, forInstructor) {
  const C = { bg: '#0b1417', panel: '#122126', line: '#27424a', ink: '#e4eee9', muted: '#8fa7a1', co2: '#f2c94c', art: '#ff7a6b', ok: '#46b97f', bad: '#ff6f5e' };
  const mono = "'Courier New',Courier,monospace";
  const sans = "Arial,Helvetica,sans-serif";
  const when = Utilities.formatDate(new Date(r.submitted), 'America/Chicago', 'MMMM d, yyyy h:mm a');
  const pct = r.max ? Math.round(100 * r.score / r.max) : 0;
  const label = t => '<div style="font-family:' + sans + ';font-size:11px;letter-spacing:2px;text-transform:uppercase;color:' + C.muted + ';font-weight:bold;">' + t + '</div>';
  const stat = (t, v, color) => '<td style="padding:0 18px 0 0;vertical-align:bottom;">' + label(t) + '<div style="font-family:' + mono + ';font-size:20px;font-weight:bold;color:' + color + ';">' + v + '</div></td>';
  const rows = r.ledger.map((l, i) => {
    const c = l.pts < 0 ? C.bad : C.ok;
    return '<tr style="background:' + (i % 2 ? C.panel : C.bg) + ';">' +
      '<td style="padding:8px;border-bottom:1px solid ' + C.line + ';font-family:' + mono + ';font-size:12px;color:' + C.muted + ';white-space:nowrap;vertical-align:top;">' + esc_(l.t) + '</td>' +
      '<td style="padding:8px;border-bottom:1px solid ' + C.line + ';font-family:' + sans + ';font-size:14px;color:' + C.ink + ';vertical-align:top;">' + esc_(l.label) + '</td>' +
      '<td style="padding:8px;border-bottom:1px solid ' + C.line + ';font-family:' + mono + ';font-size:14px;font-weight:bold;color:' + c + ';text-align:right;white-space:nowrap;vertical-align:top;">' + (l.pts > 0 ? '+' : '') + l.pts + '</td></tr>';
  }).join('');
  const expected = [
    'Recognize the early picture: rising ETCO2 despite more ventilation, tachycardia and rigidity, with FiCO2 at 0. Temperature rises late.',
    'Call for help and the MH cart, stop triggers, high flow O2 above 10 L/min, hyperventilate, charcoal filters.',
    'Dantrolene 2.5 mg/kg; older vials are 20 mg each with 60 mL sterile water.',
    'Watch for absorbent exhaustion: ETCO2 high with FiCO2 above zero means rebreathing. Change the canister.',
    'Treat acidosis and hyperkalemia; cool, and stop active cooling at 38 °C; ICU monitoring for recurrence.'
  ].map(t => '<li style="margin:0 0 6px 0;">' + t + '</li>').join('');
  const intro = forInstructor
    ? 'A student completed the scenario. This result is logged in your results sheet under the verification code below.'
    : 'Here is your report card. Your instructor received the same report. Keep this email; the verification code matches the instructor\'s record.';

  return '<!doctype html><html><body style="margin:0;padding:0;background:' + C.bg + ';">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + C.bg + ';"><tr><td align="center" style="padding:24px 12px;">' +
    '<table role="presentation" width="640" cellpadding="0" cellspacing="0" style="max-width:640px;width:100%;background:' + C.bg + ';">' +
    '<tr><td style="padding:0 0 4px 0;">' + label(COURSE + ' · Malignant hyperthermia scenario · Report card') + '</td></tr>' +
    '<tr><td style="font-family:' + sans + ';font-size:34px;font-weight:bold;color:' + C.ink + ';padding:0 0 6px 0;">OR Crisis <span style="color:' + C.art + ';">Simulator</span></td></tr>' +
    '<tr><td style="font-family:' + sans + ';font-size:15px;color:' + C.muted + ';padding:0 0 18px 0;">' + intro + '</td></tr>' +
    '<tr><td style="font-family:' + sans + ';font-size:15px;color:' + C.ink + ';padding:0 0 18px 0;"><b style="color:' + C.art + ';">Scenario revealed:</b> malignant hyperthermia</td></tr>' +
    '<tr><td style="background:' + C.panel + ';border:1px solid ' + C.line + ';border-radius:6px;padding:16px;">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' +
      '<td style="vertical-align:top;">' + label('Student') + '<div style="font-family:' + sans + ';font-size:20px;font-weight:bold;color:' + C.ink + ';">' + esc_(r.name) + '</div>' +
      '<div style="font-family:' + sans + ';font-size:13px;color:' + C.muted + ';">' + when + ' (Central)</div></td>' +
      '<td style="vertical-align:top;text-align:right;">' + label('Verification code') + '<div style="font-family:' + mono + ';font-size:20px;font-weight:bold;color:' + C.co2 + ';">' + esc_(r.code) + '</div></td>' +
      '</tr></table></td></tr>' +
    '<tr><td style="padding:20px 0 6px 0;"><table role="presentation" cellpadding="0" cellspacing="0"><tr>' +
      '<td style="padding:0 26px 0 0;vertical-align:bottom;">' + label('Final score · ' + r.max + ' possible') + '<div style="font-family:' + mono + ';font-size:54px;font-weight:bold;color:' + C.co2 + ';line-height:1;">' + r.score + '</div></td>' +
      stat('Percent', pct + '%', C.ink) + stat('Earned', '+' + r.earned, C.ok) + stat('Lost', String(r.lost), C.bad) + stat('Case time', esc_(r.caseTime), C.ink) +
    '</tr></table></td></tr>' +
    '<tr><td style="padding:16px 0;"><div style="background:' + C.panel + ';border:1px solid ' + C.line + ';border-radius:6px;padding:14px 16px;">' + label('What the expected response looks like') +
      '<ul style="margin:10px 0 0 0;padding-left:20px;font-family:' + sans + ';font-size:14px;color:' + C.ink + ';line-height:1.45;">' + expected + '</ul></div></td></tr>' +
    '<tr><td style="padding:0 0 6px 0;">' + label('Decision timeline') + '</td></tr>' +
    '<tr><td><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border-top:1px solid ' + C.line + ';">' +
      '<tr><td style="padding:8px;font-family:' + sans + ';font-size:11px;letter-spacing:2px;color:' + C.muted + ';font-weight:bold;">CASE TIME</td><td style="padding:8px;font-family:' + sans + ';font-size:11px;letter-spacing:2px;color:' + C.muted + ';font-weight:bold;">DECISION</td><td style="padding:8px;font-family:' + sans + ';font-size:11px;letter-spacing:2px;color:' + C.muted + ';font-weight:bold;text-align:right;">POINTS</td></tr>' +
      rows + '</table></td></tr>' +
    '<tr><td style="font-family:' + sans + ';font-size:12px;color:' + C.muted + ';padding:20px 0 0 0;">Oklahoma City Community College · Anesthesia Technology · ' + SCENARIO + '</td></tr>' +
    '</table></td></tr></table></body></html>';
}
