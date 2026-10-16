// Vercel serverless function: saves a visitor gate pass to OneDrive (no sign-in at the gate).
//   1. Uploads the pass PDF to   <BASE_FOLDER>/PDF/<year>/<MM-Mon>/<PassNo>_<Visitor>.pdf
//   2. Adds one row to           <BASE_FOLDER>/Visitor_Log.xlsx   (table "VisitorLog")
//      The Excel log is created automatically the first time.
//
// Uses the existing "gatepass" Entra app (same as the old gate pass site).
// Environment variables (Vercel > visitor-gatepass > Settings > Environment Variables):
//   TENANT_ID       Microsoft tenant ID
//   CLIENT_ID       "gatepass" app (client) ID
//   CLIENT_SECRET   secret VALUE of the "gatepass" app
//   ONEDRIVE_USER   ai@bharatsteels.in   (whose OneDrive receives the files)
//   BASE_FOLDER     optional, defaults to "Visitor Gate Pass"
import { LOG_TEMPLATE_B64 } from './_log-template.js';

const GRAPH = process.env.GRAPH_BASE || 'https://graph.microsoft.com/v1.0';
const LOGIN = process.env.LOGIN_BASE || 'https://login.microsoftonline.com';
const ALLOWED_ORIGINS = ['https://visitor-gatepass.vercel.app', 'https://bsc23609.github.io'];
const MAX_PDF_BYTES = 4 * 1024 * 1024;
const TABLE = 'VisitorLog';
const SHEET = 'Visitor Log';
const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

let cached = { token: null, expires: 0 };
async function getToken() {
  if (cached.token && Date.now() < cached.expires - 60000) return cached.token;
  const body = new URLSearchParams({
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });
  const r = await fetch(`${LOGIN}/${process.env.TENANT_ID}/oauth2/v2.0/token`, { method: 'POST', body });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error('Microsoft sign-in failed: ' + (j.error_description || j.error || r.status));
  cached = { token: j.access_token, expires: Date.now() + (j.expires_in || 3600) * 1000 };
  return cached.token;
}

async function graph(method, path, { json, body, contentType } = {}) {
  const token = await getToken();
  const headers = { Authorization: 'Bearer ' + token };
  if (json !== undefined) { headers['Content-Type'] = 'application/json'; body = JSON.stringify(json); }
  else if (contentType) headers['Content-Type'] = contentType;
  const r = await fetch(GRAPH + path, { method, headers, body });
  const data = await r.json().catch(() => ({}));
  return { ok: r.ok, status: r.status, data };
}
const graphError = (what, r) => new Error(`${what} (${r.status}${r.data && r.data.error ? ': ' + r.data.error.message : ''})`);

function setCors(req, res) {
  const origin = req.headers.origin || '';
  if (ALLOWED_ORIGINS.includes(origin)) res.setHeader('Access-Control-Allow-Origin', origin);
  res.setHeader('Vary', 'Origin');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

const text = (v, max) => (typeof v === 'string' ? v.trim().slice(0, max) : '');
const drivePath = parts => parts.map(encodeURIComponent).join('/');
function istNow() {
  const p = Object.fromEntries(new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Kolkata', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false,
  }).formatToParts(new Date()).map(x => [x.type, x.value]));
  return `${p.year}-${p.month}-${p.day} ${p.hour}:${p.minute}`;
}

async function ensureLog(user, base) {
  const logPath = drivePath([base, 'Visitor_Log.xlsx']);
  const found = await graph('GET', `/users/${user}/drive/root:/${logPath}`);
  if (found.ok) return found.data;
  if (found.status !== 404) throw graphError('Could not open the Excel log', found);
  const made = await graph('PUT', `/users/${user}/drive/root:/${logPath}:/content?@microsoft.graph.conflictBehavior=fail`, {
    body: Buffer.from(LOG_TEMPLATE_B64, 'base64'),
    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  if (made.ok) return made.data;
  if (made.status === 409) {               // created a moment ago by another save
    const again = await graph('GET', `/users/${user}/drive/root:/${logPath}`);
    if (again.ok) return again.data;
  }
  throw graphError('Could not create the Excel log', made);
}

async function addLogRow(user, logItem, row) {
  const wb = `/users/${user}/drive/items/${logItem.id}/workbook`;
  // The new log starts with one empty table row: fill that first, then append after it.
  const rows = await graph('GET', `${wb}/tables/${TABLE}/rows?$top=2`);
  if (!rows.ok) throw graphError('Could not read the Excel log', rows);
  const list = rows.data.value || [];
  const firstEmpty = list.length === 1 && (list[0].values || [[]])[0].every(v => v === '' || v === null);
  const r = firstEmpty
    ? await graph('PATCH', `${wb}/worksheets/${encodeURIComponent(SHEET)}/range(address='A2:J2')`, { json: { values: [row] } })
    : await graph('POST', `${wb}/tables/${TABLE}/rows/add`, { json: { values: [row] } });
  if (!r.ok) throw graphError('Could not add the row to the Excel log', r);
}

export default async function handler(req, res) {
  setCors(req, res);
  if (req.method === 'OPTIONS') return res.status(204).end();
  if (req.method === 'GET') return res.status(200).json({ ok: true, service: 'visitor-gatepass-onedrive' });
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Use POST' });

  const missingEnv = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET', 'ONEDRIVE_USER'].filter(k => !process.env[k]);
  if (missingEnv.length) return res.status(500).json({ ok: false, error: 'Missing in Vercel settings: ' + missingEnv.join(', ') });

  let b = req.body;
  if (typeof b === 'string') { try { b = JSON.parse(b); } catch { b = null; } }
  if (!b || typeof b !== 'object') return res.status(400).json({ ok: false, error: 'Invalid data' });

  const p = {
    pass_no: text(b.pass_no, 40), visitor_name: text(b.visitor_name, 60), person_to_meet: text(b.person_to_meet, 60),
    department: text(b.department, 60), visit_date: text(b.visit_date, 10), entry_time: text(b.entry_time, 8),
    purpose: text(b.purpose, 200), company: text(b.company, 120),
  };
  if (!/^VP-\d{6}-\d{6}$/.test(p.pass_no)) return res.status(400).json({ ok: false, error: 'Bad pass number' });
  if (!/^\d{4}-\d{2}-\d{2}$/.test(p.visit_date)) return res.status(400).json({ ok: false, error: 'Bad visit date' });
  if (!p.visitor_name) return res.status(400).json({ ok: false, error: 'Visitor name missing' });

  const pdf = Buffer.from(typeof b.pdfBase64 === 'string' ? b.pdfBase64 : '', 'base64');
  if (pdf.length < 100 || pdf.length > MAX_PDF_BYTES || pdf.subarray(0, 5).toString() !== '%PDF-')
    return res.status(400).json({ ok: false, error: 'Not a valid PDF' });

  const user = encodeURIComponent(process.env.ONEDRIVE_USER);
  const base = process.env.BASE_FOLDER || 'Visitor Gate Pass';
  const [y, m] = p.visit_date.split('-');
  const monthFolder = `${m}-${MONTHS[+m - 1]}`;
  const safeName = p.visitor_name.replace(/[^A-Za-z0-9]+/g, '_').replace(/^_|_$/g, '').slice(0, 40) || 'Visitor';
  const fileName = `${p.pass_no}_${safeName}.pdf`;

  let pdfUrl = '';
  try {
    const up = await graph('PUT', `/users/${user}/drive/root:/${drivePath([base, 'PDF', y, monthFolder, fileName])}:/content`, {
      body: pdf, contentType: 'application/pdf',
    });
    if (!up.ok) throw graphError('Could not save the PDF to OneDrive', up);
    pdfUrl = up.data.webUrl || '';
  } catch (e) {
    return res.status(502).json({ ok: false, error: String(e.message || e) });
  }

  try {
    const log = await ensureLog(user, base);
    const link = pdfUrl ? `=HYPERLINK("${pdfUrl.replace(/"/g, '""')}","Open PDF")` : '';
    await addLogRow(user, log, [
      p.pass_no, p.visit_date, p.entry_time, p.visitor_name, p.person_to_meet,
      p.department, p.purpose, p.company, link, istNow(),
    ]);
    return res.status(200).json({ ok: true, pdf: `${base}/PDF/${y}/${monthFolder}/${fileName}`, pdfUrl, logUrl: log.webUrl || '' });
  } catch (e) {
    return res.status(502).json({ ok: false, pdfSaved: true, pdfUrl, error: 'PDF saved, but ' + String(e.message || e) });
  }
}
