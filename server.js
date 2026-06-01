require('dotenv').config();
const express  = require('express');
const { MongoClient } = require('mongodb');
const jwt      = require('jsonwebtoken');
const XLSX     = require('xlsx');
const PDFDoc   = require('pdfkit');
const cors     = require('cors');
const path     = require('path');

const app = express();

// ── CORS ─────────────────────────────────────────────────────────────────────
const allowedOrigins = process.env.ALLOWED_ORIGINS
  ? process.env.ALLOWED_ORIGINS.split(',').map(s => s.trim())
  : null;

app.use(cors({
  origin: allowedOrigins || '*',
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));
app.use(express.json({ limit: '5mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ── MongoDB ───────────────────────────────────────────────────────────────────
let _db;
async function getDB() {
  if (!_db) {
    const client = new MongoClient(process.env.MONGODB_URI);
    await client.connect();
    _db = client.db(process.env.MONGODB_DB || 'dlf');
  }
  return _db;
}

// ── Auth middleware ───────────────────────────────────────────────────────────
const JWT_SECRET = process.env.JWT_SECRET || 'dlf-dev-secret-change-in-production';

function requireAdmin(req, res, next) {
  const auth = req.headers.authorization || '';
  if (!auth.startsWith('Bearer ')) return res.status(401).json({ error: 'Unauthorized' });
  try {
    req.user = jwt.verify(auth.slice(7), JWT_SECRET);
    next();
  } catch {
    res.status(401).json({ error: 'Token invalid or expired — please log in again' });
  }
}

// ── PUBLIC API ────────────────────────────────────────────────────────────────
// Both the self-assessment form (POST) and the assessor dashboard (GET ?code=)
// point to the same BACKEND constant, which should be set to:
//   https://YOUR-DEPLOYED-URL/api/dlf

app.post('/api/dlf', async (req, res) => {
  try {
    const { code, data } = req.body;
    if (!code || !data) return res.status(400).json({ success: false, error: 'Missing code or data' });

    const db = await getDB();
    await db.collection('submissions').replaceOne(
      { code },
      {
        code,
        data,
        fellow:    data.who   || '',
        role:      data.role  || '',
        district:  data.dist  || '',
        period:    data.period || '',
        timestamp: new Date()
      },
      { upsert: true }
    );

    res.json({ success: true, code });
  } catch (e) {
    console.error('POST /api/dlf', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

app.get('/api/dlf', async (req, res) => {
  try {
    const { code } = req.query;
    if (!code) return res.status(400).json({ success: false, error: 'Missing code param' });

    const db = await getDB();
    const doc = await db.collection('submissions').findOne({ code });
    if (!doc) return res.json({ success: false, error: 'Code not found' });

    res.json({ success: true, data: doc.data });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// ── ADMIN AUTH ────────────────────────────────────────────────────────────────
app.post('/api/admin/login', (req, res) => {
  const { username, password } = req.body || {};
  const ok =
    username === (process.env.ADMIN_USERNAME || 'admin') &&
    password === (process.env.ADMIN_PASSWORD || 'janmanindia');
  if (!ok) return res.status(401).json({ error: 'Invalid credentials' });

  const token = jwt.sign({ username, role: 'admin' }, JWT_SECRET, { expiresIn: '8h' });
  res.json({ token, expiresIn: 28800 });
});

// ── ADMIN DATA (all protected) ────────────────────────────────────────────────
app.get('/api/admin/entries', requireAdmin, async (req, res) => {
  try {
    const db = await getDB();
    const entries = await db.collection('submissions')
      .find({}, { projection: { _id: 0 } })
      .sort({ timestamp: -1 })
      .toArray();
    res.json({ success: true, count: entries.length, entries });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// ── EXCEL EXPORT ──────────────────────────────────────────────────────────────
app.get('/api/admin/export/excel', requireAdmin, async (req, res) => {
  try {
    const db = await getDB();
    const entries = await db.collection('submissions').find({}).sort({ timestamp: -1 }).toArray();

    const rows = entries.map(e => {
      const d = e.data || {};
      const scores = d.scores || {};
      const scoreVals = Object.values(scores).map(Number);
      const avg = scoreVals.length
        ? (scoreVals.reduce((a, b) => a + b, 0) / scoreVals.length).toFixed(1)
        : '';

      const row = {
        'Code':        e.code || '',
        'Timestamp':   e.timestamp ? new Date(e.timestamp).toLocaleString('en-IN') : '',
        'Fellow':      d.who    || '',
        'Role':        d.role   || '',
        'District':    d.dist   || '',
        'Period':      d.period || '',
        'Avg Score':   avg,
      };

      // Individual scores
      Object.entries(scores).forEach(([k, v]) => { row[`Score · ${k}`] = v; });

      // Section C fields
      const cMap = {
        report:   'Reporting bottleneck',
        strategy: 'Case strategy',
        media:    'Media screenings — outcomes',
        neto:     'Networking — organisations',
        netg:     'Networking — officials/govt',
        comm:     'Community work',
        sugg:     'Suggestions',
      };
      Object.entries(cMap).forEach(([k, label]) => { if (d[k]) row[label] = d[k]; });

      // Qualitative dims
      Object.entries(d.dims || {}).forEach(([k, v]) => { if (v) row[`Reflection · ${k}`] = v; });

      // Financials
      if (d.adv)   row['Advance (₹)']   = d.adv;
      if (d.claim) row['Claimed (₹)']   = d.claim;
      if (d.pend)  row['Bills pending'] = d.pend;

      // Stakeholders
      if (Array.isArray(d.stake) && d.stake.length) row['Stakeholders'] = d.stake.join(', ');

      return row;
    });

    const wb  = XLSX.utils.book_new();
    const ws  = XLSX.utils.json_to_sheet(rows);

    // Auto column widths
    if (rows.length) {
      const keys = Object.keys(rows[0]);
      ws['!cols'] = keys.map(k => ({
        wch: Math.min(50, Math.max(k.length,
          ...rows.map(r => String(r[k] || '').length).slice(0, 20)))
      }));
    }

    XLSX.utils.book_append_sheet(wb, ws, 'Submissions');
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    res.setHeader('Content-Disposition', 'attachment; filename="DLF_Submissions.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── PDF EXPORT ────────────────────────────────────────────────────────────────
app.get('/api/admin/export/pdf', requireAdmin, async (req, res) => {
  try {
    const db = await getDB();
    const entries = await db.collection('submissions').find({}).sort({ timestamp: -1 }).toArray();

    res.setHeader('Content-Disposition', 'attachment; filename="DLF_Submissions.pdf"');
    res.setHeader('Content-Type', 'application/pdf');

    const doc = new PDFDoc({ margin: 50, size: 'A4' });
    doc.pipe(res);

    // ── Cover page ──
    doc.fontSize(22).font('Helvetica-Bold').fillColor('#1a5c2e')
       .text('District Legal Fellowship', { align: 'center' });
    doc.fontSize(15).font('Helvetica').fillColor('#333')
       .text('Self-Assessment Report 2026', { align: 'center' });
    doc.fontSize(11).text("Jan Nyaya Abhiyan · Janman People's Foundation", { align: 'center' });
    doc.moveDown(0.5);
    doc.fontSize(10).fillColor('#666')
       .text(`Generated: ${new Date().toLocaleString('en-IN')}  ·  Total entries: ${entries.length}`, { align: 'center' });

    // ── Entries ──
    entries.forEach((e, idx) => {
      doc.addPage();
      const d = e.data || {};
      const scores  = d.scores  || {};
      const dims    = d.dims    || {};

      // Header
      doc.fontSize(16).font('Helvetica-Bold').fillColor('#1a5c2e')
         .text(`${idx + 1}. ${d.who || 'Unknown Fellow'}`);
      doc.fontSize(10).font('Helvetica').fillColor('#444')
         .text(`Code: ${e.code}  ·  Role: ${d.role || '—'}  ·  District: ${d.dist || '—'}  ·  Period: ${d.period || '—'}`);
      doc.text(`Submitted: ${e.timestamp ? new Date(e.timestamp).toLocaleString('en-IN') : 'Unknown'}`);
      doc.moveDown();

      // Section A — Scores
      const scoreEntries = Object.entries(scores);
      if (scoreEntries.length) {
        doc.fontSize(12).font('Helvetica-Bold').fillColor('#1a5c2e').text('Section A — Scores');
        doc.fontSize(10).font('Helvetica').fillColor('#333');
        scoreEntries.forEach(([k, v]) => {
          const n = Number(v);
          const bar = '█'.repeat(n) + '░'.repeat(10 - n);
          doc.text(`  ${String(k).padEnd(24)} ${String(v).padStart(2)}/10  ${bar}`);
        });
        const avg = (scoreEntries.reduce((a, [, v]) => a + Number(v), 0) / scoreEntries.length).toFixed(1);
        doc.font('Helvetica-Bold').text(`  ${'Average'.padEnd(24)} ${avg}/10`);
        doc.moveDown();
      }

      // Section B — Reflections
      const dimLabels = {
        experience: 'Experience',
        challenges: 'Challenges & How Overcome',
        learnings:  'Key Learnings',
        impact:     'Notable Impact',
        teamwork:   'Teamwork & Conduct'
      };
      const hasDims = Object.values(dims).some(Boolean);
      if (hasDims) {
        doc.fontSize(12).font('Helvetica-Bold').fillColor('#1a5c2e').text('Section B — Reflections');
        Object.entries(dims).forEach(([k, v]) => {
          if (!v) return;
          doc.fontSize(10).font('Helvetica-Bold').fillColor('#333')
             .text(`  ${dimLabels[k] || k}:`);
          doc.fontSize(9).font('Helvetica').fillColor('#444')
             .text(`  ${v}`, { indent: 15 });
          doc.moveDown(0.4);
        });
        doc.moveDown(0.5);
      }

      // Section C — Working the fellowship
      const cFields = [
        ['report',   'Reporting bottleneck'],
        ['strategy', 'Case strategy'],
        ['media',    'Media screenings — outcomes'],
        ['neto',     'Networking — organisations'],
        ['netg',     'Networking — officials/govt'],
        ['comm',     'Community work'],
      ];
      const hasCFields = cFields.some(([k]) => d[k]);
      if (hasCFields) {
        doc.fontSize(12).font('Helvetica-Bold').fillColor('#1a5c2e').text('Section C — Working the Fellowship');
        cFields.forEach(([k, label]) => {
          if (!d[k]) return;
          doc.fontSize(10).font('Helvetica-Bold').fillColor('#333').text(`  ${label}:`);
          doc.fontSize(9).font('Helvetica').fillColor('#444').text(`  ${d[k]}`, { indent: 15 });
          doc.moveDown(0.4);
        });
        doc.moveDown(0.5);
      }

      // Section D — Expenditure
      if (d.adv || d.claim || d.pend) {
        doc.fontSize(12).font('Helvetica-Bold').fillColor('#1a5c2e').text('Section D — Expenditure');
        doc.fontSize(10).font('Helvetica').fillColor('#333');
        if (d.adv)   doc.text(`  Advance taken: ₹${d.adv}`);
        if (d.claim) doc.text(`  Total claimed: ₹${d.claim}`);
        if (d.pend)  doc.text(`  Bills pending: ${d.pend}`);
        doc.moveDown(0.5);
      }

      // Suggestions
      if (d.sugg) {
        doc.fontSize(12).font('Helvetica-Bold').fillColor('#1a5c2e').text('Suggestions');
        doc.fontSize(9).font('Helvetica').fillColor('#444').text(`  ${d.sugg}`, { indent: 15 });
      }

      // Stakeholders
      if (Array.isArray(d.stake) && d.stake.length) {
        doc.moveDown(0.5);
        doc.fontSize(10).font('Helvetica-Bold').fillColor('#333').text('Stakeholders engaged:');
        doc.fontSize(9).font('Helvetica').text(`  ${d.stake.join(', ')}`);
      }
    });

    doc.end();
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── ADMIN PAGE ────────────────────────────────────────────────────────────────
app.get('/admin', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'admin.html'));
});

// Health check
app.get('/', (req, res) => res.json({ status: 'ok', service: 'DLF API', version: '2.0' }));

// ── START ─────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`DLF API running → http://localhost:${PORT}`));
module.exports = app; // for Vercel
