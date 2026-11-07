/**
 * RSVP backend for https://joyceetradvin.github.io
 * Standalone script; creates/uses the "RSVP — Joyce & Radvin" Google Sheet (ID kept in Script Properties).
 * Deploy: Deploy > New deployment > Web app · Execute as: Me · Who has access: Anyone
 */
const RECIPIENTS = 'ravinderupem@gmail.com,joycedjian@gmail.com';
const HEADERS = ['Date', 'Nom', 'Email', 'Présence', 'Message'];

function ss_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const ss = SpreadsheetApp.create('RSVP — Joyce & Radvin');
  ss.getSheets()[0].appendRow(HEADERS);
  ss.getSheets()[0].setFrozenRows(1);
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

/** Run once from the editor to create the Sheet and authorize. */
function setup() {
  Logger.log(ss_().getUrl());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d._honey) return json_({ ok: true }); // bot

    const clip = (v, n) => String(v || '').trim().slice(0, n);
    const nom = clip(d.nom, 120);
    const email = clip(d.email, 160);
    const presence = clip(d.presence, 80);
    const message = clip(d.message, 3000);
    if (!nom || !email || !presence) return json_({ ok: false, error: 'missing fields' });

    const ss = ss_();
    const sh = ss.getSheets()[0];
    if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
    sh.appendRow([new Date(), nom, email, presence, message]);

    const esc = (s) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    const row = (k, v) => '<tr><td style="padding:6px 12px;color:#666">' + k + '</td><td style="padding:6px 12px"><b>' + esc(v || '—').replace(/\n/g, '<br>') + '</b></td></tr>';
    const html = '<p>Nouvelle réponse RSVP :</p><table style="border-collapse:collapse;font-family:sans-serif">'
      + row('Nom', nom) + row('Email', email) + row('Présence', presence) + row('Message', message)
      + '</table><p style="color:#999;font-size:12px">Toutes les réponses : ' + ss.getUrl() + '</p>';

    const opts = { to: RECIPIENTS, subject: 'RSVP — ' + nom + ' — ' + presence, htmlBody: html, name: 'RSVP Joyce & Radvin' };
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) opts.replyTo = email;
    MailApp.sendEmail(opts);

    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, service: 'rsvp' });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
