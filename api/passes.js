// Vercel serverless function: saves each submitted visitor gate pass into Neon Postgres.
// Needs one environment variable in Vercel: DATABASE_URL (the Neon connection string).
import pg from 'pg';

const ALLOWED_ORIGINS = [
  'https://visitor-gatepass.vercel.app',
  'https://bsc23609.github.io',
];

const MAX_IMAGE_CHARS = 1_500_000; // ~1.1 MB per image, far above a normal pass photo

let pool;
function getPool() {
  if (!pool) {
    pool = new pg.Pool({
      connectionString: process.env.DATABASE_URL,
      ssl: process.env.DATABASE_URL && process.env.DATABASE_URL.includes('localhost') ? false : { rejectUnauthorized: false },
      max: 1,
    });
  }
  return pool;
}

function setCors(req, res) {
  const origin = req.headers.origin;
  if (ALLOWED_ORIGINS.includes(origin)) res.setHeader('Access-Control-Allow-Origin', origin);
  res.setHeader('Vary', 'Origin');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

const text = (v, max) => (typeof v === 'string' ? v.trim().slice(0, max) : '');
function image(v, kinds) {
  if (typeof v !== 'string' || !v) return null;
  if (v.length > MAX_IMAGE_CHARS) return null;
  return kinds.some(k => v.startsWith(`data:image/${k};base64,`)) ? v : null;
}

export default async function handler(req, res) {
  setCors(req, res);
  if (req.method === 'OPTIONS') return res.status(204).end();
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Use POST' });

  if (!process.env.DATABASE_URL) {
    return res.status(500).json({ ok: false, error: 'DATABASE_URL is not set in Vercel' });
  }

  let body = req.body;
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch { body = null; } }
  if (!body || typeof body !== 'object') return res.status(400).json({ ok: false, error: 'Invalid data' });

  const rec = {
    pass_no:        text(body.pass_no, 40),
    company:        text(body.company, 120),
    visitor_name:   text(body.visitor_name, 60),
    person_to_meet: text(body.person_to_meet, 60),
    department:     text(body.department, 60),
    visit_date:     text(body.visit_date, 10),
    entry_time:     text(body.entry_time, 8),
    purpose:        text(body.purpose, 200),
    photo:          image(body.photo, ['jpeg', 'png']),
    signature:      image(body.security_signature, ['png']),
  };

  const missing = ['pass_no', 'visitor_name', 'person_to_meet', 'department', 'visit_date', 'entry_time', 'purpose']
    .filter(k => !rec[k]);
  if (missing.length) return res.status(400).json({ ok: false, error: 'Missing: ' + missing.join(', ') });
  if (!/^\d{4}-\d{2}-\d{2}$/.test(rec.visit_date)) return res.status(400).json({ ok: false, error: 'Bad visit_date' });
  if (!/^\d{2}:\d{2}(:\d{2})?$/.test(rec.entry_time)) return res.status(400).json({ ok: false, error: 'Bad entry_time' });

  try {
    const r = await getPool().query(
      `INSERT INTO visitor_passes
         (pass_no, company, visitor_name, person_to_meet, department, visit_date, entry_time, purpose, photo, security_signature)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10)
       ON CONFLICT (pass_no) DO UPDATE SET pass_no = EXCLUDED.pass_no
       RETURNING id, created_at`,
      [rec.pass_no, rec.company, rec.visitor_name, rec.person_to_meet, rec.department,
       rec.visit_date, rec.entry_time, rec.purpose, rec.photo, rec.signature]
    );
    return res.status(200).json({ ok: true, id: r.rows[0].id, saved_at: r.rows[0].created_at });
  } catch (err) {
    console.error('save failed', err);
    return res.status(500).json({ ok: false, error: 'Database error' });
  }
}
