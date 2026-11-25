/**
 * Bluff Road Bike Co — /waiver backend
 * Bind this to the "BRBC Waiver" Google Sheet (levi@herdcoindy.com):
 *   Extensions → Apps Script → paste this file → Deploy → New deployment →
 *   Web app · Execute as: Me · Who has access: Anyone → copy the /exec URL.
 *
 * Each signature:
 *   1. appends a row to the "Waivers" tab (same columns the Google Form used, so
 *      Completed / Brand / Style and the about-page stats keep working),
 *   2. saves a signed PDF (full agreement text + signature + audit trail) to a
 *      Drive folder "BRBC Signed Waivers",
 *   3. emails the PDF to the customer and a heads-up to the shop.
 */

const SHOP_NAME   = 'Bluff Road Bike Co LLC';
const SHOP_EMAIL  = 'levi@bluffroadbikeco.com';   // reply-to + notification inbox
const SHEET_TAB   = 'Waivers';
const FOLDER_NAME = 'BRBC Signed Waivers';
const TZ          = 'America/Indiana/Indianapolis';

// Header names (existing columns are matched by name; new ones are added at the end)
const H = {
  ts: 'Timestamp',
  ack: 'CUSTOMER ACKNOWLEDGMENT\nBy signing below, I confirm that I have read and agree to the terms of this Release of Liability & Service Waiver.',
  sigName: 'Type Name as Signature',
  phone: 'Phone Number',
  email: 'Email Address',
  date: 'Date of Service',
  makeModel: 'Bike Make/Model',
  notes: 'Notes',
  completed: 'Completed',
  brand: 'Brand',
  style: 'Style',
  ref: 'Ref ID',
  version: 'Waiver Version',
  initials: 'Initials',
  minor: 'Minor Rider',
  relationship: 'Guardian Relationship',
  pdf: 'Signed PDF',
  hash: 'Terms SHA-256',
  client: 'Client Time / Zone',
  ua: 'User Agent',
  source: 'Source'
};

function doGet() {
  return json_({ ok: true, service: 'BRBC waiver' });
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.company) return json_({ ok: true, ref: 'BRBC-OK' });   // honeypot: quietly drop bots

    // ── validate ──
    const need = ['name', 'email', 'phone', 'serviceDate', 'style', 'brand', 'typedName', 'signaturePng', 'termsText', 'initials'];
    for (const k of need) if (!d[k] || String(d[k]).trim() === '') throw new Error('Missing ' + k + '.');
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email)) throw new Error('Invalid email.');
    if (norm_(d.typedName) !== norm_(d.name)) throw new Error('Typed name does not match.');
    if (!/^data:image\/png;base64,/.test(d.signaturePng) || d.signaturePng.length > 600000) throw new Error('Invalid signature image.');
    if (d.termsText.length > 20000) throw new Error('Invalid agreement text.');
    if (d.isMinor && (!d.minorName || !d.relationship)) throw new Error('Minor rider details missing.');

    const now = new Date();
    const serverHash = sha256_(d.termsText);
    const ref = 'BRBC-' + Utilities.formatDate(now, TZ, 'yyMMdd') + '-' +
                Utilities.getUuid().replace(/-/g, '').slice(0, 5).toUpperCase();

    // ── PDF ──
    const pdf = buildPdf_(d, ref, now, serverHash);
    const file = folder_().createFile(pdf);

    // ── Sheet row ──
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_TAB);
      const col = ensureHeaders_(sh);
      const row = new Array(sh.getLastColumn()).fill('');
      const set = (key, val) => { row[col[H[key]] - 1] = val; };
      set('ts', now);
      set('ack', 'I Agree');
      set('sigName', d.typedName);
      set('phone', d.phone);
      set('email', d.email);
      set('date', d.serviceDate);
      set('makeModel', [d.brand, d.model].filter(Boolean).join(' '));
      set('notes', d.notes || '');
      set('completed', false);
      set('brand', d.brand);
      set('style', d.style);
      set('ref', ref);
      set('version', d.waiverVersion || '');
      set('initials', d.initials);
      set('minor', d.isMinor ? d.minorName : '');
      set('relationship', d.isMinor ? d.relationship : '');
      set('pdf', file.getUrl());
      set('hash', serverHash);
      set('client', (d.signedAtClient || '') + ' / ' + (d.timezone || ''));
      set('ua', (d.userAgent || '').slice(0, 300));
      set('source', 'bluffroadbikeco.com/waiver');
      sh.appendRow(row);
      // keep the Completed column a checkbox
      sh.getRange(sh.getLastRow(), col[H.completed]).insertCheckboxes();
    } finally {
      lock.releaseLock();
    }

    // ── Emails ──
    sendMail_(d.email,
      'Your signed service agreement — ' + SHOP_NAME + ' (' + ref + ')',
      'Hi ' + d.name.split(' ')[0] + ',\n\n' +
      'Thanks for signing the Bluff Road Bike Co Release of Liability & Service Agreement. ' +
      'Your signed copy is attached for your records.\n\n' +
      'Reference: ' + ref + '\nDrop-off date: ' + d.serviceDate + '\nBike: ' + [d.brand, d.model].filter(Boolean).join(' ') + '\n\n' +
      'Need a drop-off time? https://calendly.com/bluffroadbikeco\n\n— Levi\nBluff Road Bike Co\n(317) 721-7339',
      [pdf]);

    sendMail_(SHOP_EMAIL,
      'New waiver: ' + d.name + ' · ' + d.brand + ' ' + d.style + ' · ' + d.serviceDate,
      d.name + ' signed waiver ' + ref + '.\n\nPhone: ' + d.phone + '\nEmail: ' + d.email +
      '\nBike: ' + [d.brand, d.model].filter(Boolean).join(' ') + ' (' + d.style + ')' +
      (d.isMinor ? '\nMinor rider: ' + d.minorName + ' (' + d.relationship + ' signed)' : '') +
      '\n\nNotes: ' + (d.notes || '—') + '\n\nPDF: ' + file.getUrl(),
      []);

    return json_({ ok: true, ref: ref });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

// ───────── helpers ─────────

function buildPdf_(d, ref, now, hash) {
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const signedAt = Utilities.formatDate(now, TZ, "MMMM d, yyyy 'at' h:mm:ss a z");
  const terms = d.termsText.split('\n\n').map(block => {
    const lines = block.split('\n');
    return '<h3>' + esc(lines[0]) + '</h3>' + lines.slice(1).map(p => '<p>' + esc(p) + '</p>').join('');
  }).join('');
  const bike = [d.brand, d.model].filter(Boolean).join(' ') + ' (' + d.style + ')';

  const html = `
  <html><head><style>
    body { font-family: Arial, Helvetica, sans-serif; color:#1b2433; font-size:10.5pt; line-height:1.45; }
    .hdr { border-bottom:3px solid #021328; padding-bottom:8px; margin-bottom:6px; }
    .hdr b { font-size:16pt; color:#021328; letter-spacing:1px; }
    .hdr div { font-size:9pt; color:#5b6472; }
    h1 { font-size:15pt; text-align:center; color:#021328; margin:14px 0 2px; }
    .meta { text-align:center; font-size:9pt; color:#5b6472; margin-bottom:10px; }
    h3 { font-size:10.5pt; color:#021328; margin:10px 0 2px; text-transform:uppercase; }
    p { margin:0 0 4px; }
    table { width:100%; border-collapse:collapse; margin-top:8px; }
    td { border:1px solid #ccc; padding:5px 7px; vertical-align:top; font-size:10pt; }
    td.k { width:32%; background:#f5f0e4; font-weight:bold; }
    .sig { border-bottom:1px solid #333; padding:4px 0; }
    .sig img { height:70px; }
    .audit td { font-size:8.5pt; }
    .sec { font-size:12pt; color:#021328; border-bottom:2px solid #021328; margin-top:16px; padding-bottom:2px; }
  </style></head><body>
    <div class="hdr"><b>BLUFF ROAD BIKE CO LLC</b><div>Bicycle Service · Greenwood, Indiana · ${esc(SHOP_EMAIL)}</div></div>
    <h1>Release of Liability &amp; Service Agreement</h1>
    <div class="meta">Version ${esc(d.waiverVersion)} · Reference ${esc(ref)}</div>
    ${terms}

    <div class="sec">Customer &amp; Service</div>
    <table>
      <tr><td class="k">Customer</td><td>${esc(d.name)}</td></tr>
      <tr><td class="k">Email / Phone</td><td>${esc(d.email)} · ${esc(d.phone)}</td></tr>
      <tr><td class="k">Drop-off date</td><td>${esc(d.serviceDate)}</td></tr>
      <tr><td class="k">Bicycle</td><td>${esc(bike)}</td></tr>
      ${d.notes ? `<tr><td class="k">Requested work</td><td>${esc(d.notes)}</td></tr>` : ''}
      ${d.isMinor ? `<tr><td class="k">Minor rider</td><td>${esc(d.minorName)} — signed by ${esc(d.relationship.toLowerCase())}</td></tr>` : ''}
      <tr><td class="k">Initials (§3, §4, §6)</td><td>${esc(d.initials)} · ${esc(d.initials)} · ${esc(d.initials)}</td></tr>
    </table>

    <div class="sec">Signature</div>
    <table>
      <tr><td class="k">Signature</td><td><div class="sig"><img src="${d.signaturePng}"></div></td></tr>
      <tr><td class="k">Printed name</td><td>${esc(d.typedName)}</td></tr>
      <tr><td class="k">Signed</td><td>${esc(signedAt)}</td></tr>
    </table>

    <div class="sec">Electronic Signature Record</div>
    <table class="audit">
      <tr><td class="k">Consents</td><td>Read &amp; agree: yes · Electronic signature &amp; records (E-SIGN / IC 26-2-8): yes</td></tr>
      <tr><td class="k">Server timestamp</td><td>${esc(now.toISOString())}</td></tr>
      <tr><td class="k">Device time / zone</td><td>${esc(d.signedAtClient)} / ${esc(d.timezone)}</td></tr>
      <tr><td class="k">Agreement SHA-256</td><td style="word-break:break-all">${esc(hash)}</td></tr>
      <tr><td class="k">Signed at</td><td>${esc(d.pageUrl)}</td></tr>
      <tr><td class="k">Device</td><td>${esc((d.userAgent || '').slice(0, 300))}</td></tr>
    </table>
  </body></html>`;

  const safeName = d.name.replace(/[^\w\- ]/g, '').trim();
  return Utilities.newBlob(html, 'text/html', 'tmp.html')
    .getAs('application/pdf')
    .setName(`${ref} - ${safeName} - BRBC Waiver.pdf`);
}

function ensureHeaders_(sh) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const map = {};
  headers.forEach((h, i) => { if (h) map[h] = i + 1; });
  let next = lastCol + 1;
  Object.values(H).forEach(h => {
    if (!map[h]) {
      // the long Form acknowledgment header may differ slightly — match by prefix
      const hit = headers.findIndex(x => x.indexOf('CUSTOMER ACKNOWLEDGMENT') === 0);
      if (h === H.ack && hit >= 0) { map[h] = hit + 1; return; }
      if (sh.getMaxColumns() < next) sh.insertColumnsAfter(sh.getMaxColumns(), 1);
      sh.getRange(1, next).setValue(h).setFontWeight('bold');
      map[h] = next++;
    }
  });
  return map;
}

function folder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  const f = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}

function sendMail_(to, subject, body, attachments) {
  const opts = { name: 'Bluff Road Bike Co', replyTo: SHOP_EMAIL, attachments: attachments };
  if (GmailApp.getAliases().indexOf(SHOP_EMAIL) >= 0) opts.from = SHOP_EMAIL;
  GmailApp.sendEmail(to, subject, body, opts);
}

function sha256_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}

function norm_(s) { return String(s).trim().replace(/\s+/g, ' ').toLowerCase(); }

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Run once from the editor to grant permissions and test PDF + email to yourself. */
function testSignature() {
  const fake = {
    postData: { contents: JSON.stringify({
      waiverVersion: '2026.11', termsText: '1. Test\nThis is a test agreement.', initials: 'LM',
      name: 'Test Customer', email: SHOP_EMAIL, phone: '317-555-0100', serviceDate: '2026-11-25',
      style: 'Road', brand: 'Test', model: 'Delete me', notes: 'Test row — delete', isMinor: false,
      typedName: 'Test Customer',
      signaturePng: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
      signedAtClient: new Date().toISOString(), timezone: TZ, userAgent: 'Apps Script test', pageUrl: 'test'
    }) }
  };
  Logger.log(doPost(fake).getContent());
}
