// ══════════════════════════════════════════════════════════════════════════════
//  KIU CBE Programme Generator — Backend API v4.0
//  Storage: Google Sheets (replaces PostgreSQL)
//  All API endpoints identical to v3 — index.html needs NO changes.
// ══════════════════════════════════════════════════════════════════════════════

const express    = require('express');
const cors       = require('cors');
const fetch      = require('node-fetch');
const { google } = require('googleapis');

// ── Resilient fetch — retries transient network failures ──────────────────────
// "Premature close", socket hang-ups, ECONNRESET and timeouts from the Render
// instance to external APIs (OpenRouter, Google) are retried automatically.
async function fetchWithRetry(url, options, maxAttempts = 3, baseDelayMs = 1500) {
  let lastErr;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const r = await fetch(url, options);
      return r; // success (any HTTP status — caller decides what to do with non-2xx)
    } catch (e) {
      lastErr = e;
      const msg = (e && e.message) || '';
      const transient = /premature close|socket hang up|ECONNRESET|ETIMEDOUT|EAI_AGAIN|network|timeout|aborted/i.test(msg)
        || e.code === 'ECONNRESET' || e.code === 'ETIMEDOUT' || e.type === 'request-timeout';
      if (transient && attempt < maxAttempts) {
        const wait = baseDelayMs * attempt;
        console.log(`⏳ fetch ${url.slice(0, 60)} attempt ${attempt}/${maxAttempts} failed (${msg}) — retrying in ${wait}ms`);
        await new Promise(res => setTimeout(res, wait));
        continue;
      }
      throw e;
    }
  }
  throw lastErr;
}

const app  = express();
const PORT = process.env.PORT || 3001;

app.use(cors({ origin: '*' }));
app.use(express.json({ limit: '16mb' }));

// ── Sheet tab names ───────────────────────────────────────────────────────────
const SHEET = {
  REGISTRY:      'Registry',
  PROGRAMMES:    'Programmes',
  DETAILS:       'CourseDetails',
  SECTIONS:      'ProgrammeSections',
  NOTIFICATIONS: 'Notifications',
  BORROWS:       'Borrows',
};

// ── Column headers — order matters, maps 1:1 to Sheets columns ───────────────
const HEADERS = {
  REGISTRY: [
    'id','code','name','prefix','programme','school','dept',
    'tier','dh','sh','ah','oh','type','year_num','sem',
    'developer','owner_school','owner_dept','updated_at',
  ],
  PROGRAMMES: [
    'id','name','abbr','school','dept',
    'meta','courses','section_done','course_done',
    'course_count','section_content','saved_at',
  ],
  DETAILS: [
    'id','code','name','programme','school','dept',
    'tier','content','updated_at',
  ],
  SECTIONS: [
    'id','programme_id','programme_name',
    'section_key','content','done','updated_at',
  ],
  NOTIFICATIONS: [
    'id','type','title','message',
    'for_school','for_dept','from_school','from_programme',
    'developer','is_read','created_at',
  ],
  BORROWS: [
    'id','course_code','course_name',
    'borrower_prog','borrower_school','borrower_dept',
    'owner_prog','owner_school','owner_dept','owner_code',
    'developer','decision','position_adopted','created_at',
  ],
};

// Max chars to store in a cell (Sheets hard limit is 50,000).
// Large text (course outlines, JSON blobs) is truncated with a warning marker.
const CELL_LIMIT = 48000;

// ── Google Sheets state ───────────────────────────────────────────────────────
let sheetsClient = null;
const SPREADSHEET_ID = process.env.SPREADSHEET_ID || '';

// ── Initialise Google Sheets connection ───────────────────────────────────────
async function initSheets() {
  const credJson = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (!credJson) {
    console.warn('⚠️  GOOGLE_SERVICE_ACCOUNT_JSON not set — running without storage');
    return false;
  }
  if (!SPREADSHEET_ID) {
    console.warn('⚠️  SPREADSHEET_ID not set — running without storage');
    return false;
  }
  try {
    const credentials = JSON.parse(credJson);
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    sheetsClient = google.sheets({ version: 'v4', auth });
    await ensureAllTabs();
    console.log('✅ Google Sheets connected — all tabs ready');
    return true;
  } catch (e) {
    console.error('❌ Sheets init failed:', e.message);
    sheetsClient = null;
    return false;
  }
}

// ── Create any missing sheet tabs and write their header rows ─────────────────
async function ensureAllTabs() {
  const meta = await sheetsClient.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const existing = new Set(meta.data.sheets.map(s => s.properties.title));

  for (const [key, tabName] of Object.entries(SHEET)) {
    if (!existing.has(tabName)) {
      await sheetsClient.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: { requests: [{ addSheet: { properties: { title: tabName } } }] },
      });
      await sheetsClient.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range:            `${tabName}!A1`,
        valueInputOption: 'RAW',
        requestBody:      { values: [HEADERS[key]] },
      });
      console.log(`  ✅ Created tab: ${tabName}`);
    }
  }
}

// ════════════════════════════════════════════════════════════════════════════
//  GENERIC SHEETS HELPERS
// ════════════════════════════════════════════════════════════════════════════

// Read all data rows from a tab.
// Returns array of objects; each object includes _rowIndex (1-based sheet row).
async function getRows(tabName) {
  if (!sheetsClient) throw new Error('Google Sheets not connected. Check env vars on Render.');
  const res = await sheetsClient.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range:         `${tabName}!A:ZZ`,
  });
  const raw = res.data.values || [];
  if (raw.length < 2) return [];          // header only or empty
  const headers = raw[0];
  return raw.slice(1).map((row, i) => {
    const obj = { _rowIndex: i + 2 };    // +1 for header, +1 for 1-based index
    headers.forEach((h, j) => { obj[h] = row[j] !== undefined ? row[j] : ''; });
    return obj;
  });
}

// Overwrite a single row identified by its 1-based sheet row index.
async function updateRow(tabName, rowIndex, headerKey, values) {
  const hdrs = HEADERS[headerKey];
  const lastCol = String.fromCharCode(64 + hdrs.length);   // e.g. 'S' for 19 columns
  await sheetsClient.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range:            `${tabName}!A${rowIndex}:${lastCol}${rowIndex}`,
    valueInputOption: 'RAW',
    requestBody:      { values: [safeValues(values)] },
  });
}

// Append a new row at the end of the sheet.
async function appendRow(tabName, values) {
  await sheetsClient.spreadsheets.values.append({
    spreadsheetId:   SPREADSHEET_ID,
    range:           `${tabName}!A1`,
    valueInputOption: 'RAW',
    insertDataOption: 'INSERT_ROWS',
    requestBody:     { values: [safeValues(values)] },
  });
}

// Ensure every value is a string and respects the cell character limit.
function safeValues(values) {
  return values.map(v => {
    if (v === null || v === undefined) return '';
    const s = typeof v === 'object' ? JSON.stringify(v) : String(v);
    return s.length > CELL_LIMIT ? s.slice(0, CELL_LIMIT) + '…[truncated]' : s;
  });
}

// Serialise a rowData object into an ordered array matching the header definition.
function rowToValues(headerKey, rowData) {
  return HEADERS[headerKey].map(h => {
    const v = rowData[h];
    if (v === null || v === undefined) return '';
    if (typeof v === 'object') return JSON.stringify(v);
    return String(v);
  });
}

// Safely parse a JSON cell value; returns fallback on failure.
function parseCell(raw, fallback) {
  if (!raw || raw === '') return fallback;
  try { return JSON.parse(raw); } catch { return fallback; }
}

// ════════════════════════════════════════════════════════════════════════════
//  DOMAIN CONVERTERS
// ════════════════════════════════════════════════════════════════════════════

// Sheets row → registry API object
function toReg(r) {
  return {
    code: r.code, name: r.name, prefix: r.prefix,
    programme: r.programme, school: r.school, dept: r.dept,
    tier: parseInt(r.tier) || 4,
    DH: parseInt(r.dh) || 0, SH: parseInt(r.sh) || 0,
    AH: parseInt(r.ah) || 0, OH: parseInt(r.oh) || 0,
    type: r.type, year: parseInt(r.year_num) || 1, sem: r.sem,
    developer: r.developer, ownerSchool: r.owner_school,
    ownerDept: r.owner_dept, updatedAt: r.updated_at,
  };
}

// Sheets row → programme API object
function toProg(p) {
  return {
    id: p.id, name: p.name, abbr: p.abbr, school: p.school, dept: p.dept,
    meta:           parseCell(p.meta, {}),
    courses:        parseCell(p.courses, []),
    sectionDone:    parseCell(p.section_done, {}),
    courseDone:     parseCell(p.course_done, {}),
    courseCount:    parseInt(p.course_count) || 0,
    sectionContent: parseCell(p.section_content, {}),
    savedAt:        p.saved_at,
  };
}

// ════════════════════════════════════════════════════════════════════════════
//  HEALTH
// ════════════════════════════════════════════════════════════════════════════

app.get('/api/health', async (req, res) => {
  if (!sheetsClient) {
    return res.json({
      status: 'ok-no-db', db: 'not connected',
      message: 'GOOGLE_SERVICE_ACCOUNT_JSON or SPREADSHEET_ID not set',
      env: {
        SPREADSHEET_ID:               !!process.env.SPREADSHEET_ID,
        GOOGLE_SERVICE_ACCOUNT_JSON:  !!process.env.GOOGLE_SERVICE_ACCOUNT_JSON,
        OPENROUTER_API_KEY:           !!process.env.OPENROUTER_API_KEY,
      },
    });
  }
  try {
    const [reg, notifs, progs] = await Promise.all([
      getRows(SHEET.REGISTRY),
      getRows(SHEET.NOTIFICATIONS),
      getRows(SHEET.PROGRAMMES),
    ]);
    res.json({
      status: 'ok', db: 'google-sheets',
      stats: {
        totalCourses:  reg.length,
        notifications: notifs.filter(n => n.is_read !== 'true').length,
        programmes:    progs.length,
      },
    });
  } catch (e) {
    res.status(500).json({ status: 'error', error: e.message });
  }
});

// ════════════════════════════════════════════════════════════════════════════
//  AI CHAT  (unchanged — still proxies OpenRouter)
// ════════════════════════════════════════════════════════════════════════════

app.post('/api/chat', async (req, res) => {
  const { prompt, system, maxTokens } = req.body;
  if (!prompt) return res.status(400).json({ error: 'prompt required' });
  const apiKey = process.env.OPENROUTER_API_KEY;
  if (!apiKey) return res.status(500).json({ error: 'OPENROUTER_API_KEY not set on server' });
  try {
    const r = await fetchWithRetry('https://openrouter.ai/api/v1/chat/completions', {
      method: 'POST',
      timeout: 120000, // 2-min cap per attempt (node-fetch v2 honours this)
      headers: {
        'Content-Type':  'application/json',
        'Authorization': 'Bearer ' + apiKey,
        'HTTP-Referer':  'https://kiu.ac.ug',
        'X-Title':       'KIU CBE Programme Generator',
      },
      body: JSON.stringify({
        model:       'meta-llama/llama-3.3-70b-instruct',
        max_tokens:  maxTokens || 4000,
        temperature: 0.7,
        messages: [
          { role: 'system', content: system || 'You are a helpful assistant.' },
          { role: 'user',   content: prompt },
        ],
      }),
    });
    if (!r.ok) {
      const errText = await r.text();
      return res.status(r.status).json({ error: errText.slice(0, 300) });
    }
    const data = await r.json();
    res.json({ content: data.choices?.[0]?.message?.content || '', model: data.model });
  } catch (e) {
    res.status(500).json({ error: 'AI request failed after retries: ' + e.message });
  }
});

// ════════════════════════════════════════════════════════════════════════════
//  COURSE REGISTRY
// ════════════════════════════════════════════════════════════════════════════

app.get('/api/registry', async (req, res) => {
  try {
    let rows = await getRows(SHEET.REGISTRY);
    if (req.query.school) rows = rows.filter(r => r.school === req.query.school);
    if (req.query.dept)   rows = rows.filter(r => r.dept   === req.query.dept);
    if (req.query.tier)   rows = rows.filter(r => parseInt(r.tier) === parseInt(req.query.tier));
    rows.sort((a, b) => new Date(b.updated_at || 0) - new Date(a.updated_at || 0));
    res.json(rows.slice(0, 1000).map(toReg));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/registry', async (req, res) => {
  const { programme, school, dept, courses, developer } = req.body;
  if (!courses || !Array.isArray(courses))
    return res.status(400).json({ error: 'courses array required' });

  // Read once, update in-memory index to avoid N+1 reads
  const existingRows = await getRows(SHEET.REGISTRY);
  const byKey = {};
  existingRows.forEach(r => { byKey[`${r.code}||${r.programme}`] = r; });

  let saved = 0, updated = 0;

  for (const c of courses) {
    if (!c.code && !c.name) continue;
    const now = new Date().toISOString();
    const rowData = {
      id:          `${c.code || ''}||${programme || ''}`,
      code:        c.code        || '',
      name:        c.name        || '',
      prefix:      c.prefix      || '',
      programme:   programme     || '',
      school:      school || c.school || '',
      dept:        dept   || c.dept   || '',
      tier:        c.tier        || 4,
      dh:          c.DH          || 0,
      sh:          c.SH          || 0,
      ah:          c.AH          || 0,
      oh:          c.OH          || 0,
      type:        c.type        || 'lec',
      year_num:    c.year        || 1,
      sem:         c.sem         || '1',
      developer:   developer     || '',
      owner_school: c.ownerSchool || '',
      owner_dept:   c.ownerDept  || '',
      updated_at:  now,
    };
    const values = rowToValues('REGISTRY', rowData);
    const key    = `${c.code || ''}||${programme || ''}`;
    try {
      if (byKey[key]) {
        await updateRow(SHEET.REGISTRY, byKey[key]._rowIndex, 'REGISTRY', values);
        updated++;
      } else {
        await appendRow(SHEET.REGISTRY, values);
        saved++;
        byKey[key] = { ...rowData, _rowIndex: existingRows.length + 2 + saved };
      }
    } catch (e) { console.warn('Registry row error:', e.message); }
  }
  res.json({ saved, updated });
});

// Fuzzy-match course check against existing registry
app.post('/api/registry/check', async (req, res) => {
  const { name, code, school, dept, programme, year, sem } = req.body;
  if (!name && !code) return res.status(400).json({ error: 'name or code required' });
  try {
    const nameLower  = (name || '').toLowerCase().replace(/[^a-z0-9 ]/g, '').trim();
    const nameWords  = nameLower.split(/\s+/).filter(w => w.length > 3);
    const threshold  = nameWords.length <= 3 ? nameWords.length : Math.ceil(nameWords.length * 0.65);
    const allRows    = await getRows(SHEET.REGISTRY);
    const seen       = new Set();
    const candidates = [];

    if (code) {
      allRows
        .filter(r => r.code.toUpperCase() === (code || '').toUpperCase() && r.programme !== (programme || ''))
        .forEach(r => { const k = r.code + r.programme; seen.add(k); candidates.push(r); });
    }
    if (nameWords.length > 0) {
      allRows
        .filter(r => {
          if (r.programme === (programme || '')) return false;
          const rW = (r.name || '').toLowerCase().replace(/[^a-z0-9 ]/g, '').split(/\s+/).filter(w => w.length > 3);
          return nameWords.filter(w => rW.includes(w)).length >= threshold;
        })
        .forEach(r => {
          const k = r.code + r.programme;
          if (!seen.has(k)) { seen.add(k); candidates.push(r); }
        });
    }

    if (!candidates.length) return res.json({ hasConflicts: false, conflicts: [] });

    const mapped        = candidates.slice(0, 10).map(toReg);
    const isCrossSchool = mapped.some(r => r.school && r.school !== school);
    const isCrossDept   = mapped.some(r => r.dept && r.dept !== dept && r.school === school);
    res.json({
      hasConflicts: true, isCrossSchool, isCrossDept, conflicts: mapped,
      suggestedTier:    isCrossSchool ? 2 : isCrossDept ? 3 : 4,
      positionMismatch: mapped.some(r => (r.year && r.year !== year) || (r.sem && r.sem !== sem)),
    });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Record a borrow/reclassify event + post notification
app.post('/api/registry/borrow', async (req, res) => {
  const {
    courseCode, courseName, borrowerProgramme, borrowerSchool, borrowerDept,
    ownerProgramme, ownerSchool, ownerDept, ownerCode, developer, decision, positionAdopted,
  } = req.body;
  try {
    const now = new Date().toISOString();
    const borrowData = {
      id:               Date.now(),
      course_code:      courseCode      || '',
      course_name:      courseName      || '',
      borrower_prog:    borrowerProgramme || '',
      borrower_school:  borrowerSchool  || '',
      borrower_dept:    borrowerDept    || '',
      owner_prog:       ownerProgramme  || '',
      owner_school:     ownerSchool     || '',
      owner_dept:       ownerDept       || '',
      owner_code:       ownerCode       || '',
      developer:        developer       || '',
      decision:         decision        || 'borrow',
      position_adopted: !!positionAdopted,
      created_at:       now,
    };
    await appendRow(SHEET.BORROWS, rowToValues('BORROWS', borrowData));

    if (decision === 'borrow' || decision === 'reclassify') {
      const msg = decision === 'borrow'
        ? `${borrowerProgramme} is using "${courseName}" (${ownerCode || ''}) from ${ownerSchool}/${ownerDept}.${positionAdopted ? ' Position adopted.' : ''}`
        : `${borrowerProgramme} requested reclassification of "${courseName}".`;
      const notifData = {
        id:               Date.now() + 1,
        type:             decision === 'borrow' ? 'BORROW' : 'RECLASSIFY',
        title:            `"${courseName}" ${decision === 'borrow' ? 'borrowed by' : 'reclassification from'} ${borrowerSchool}`,
        message:          msg,
        for_school:       ownerSchool       || '',
        for_dept:         ownerDept         || '',
        from_school:      borrowerSchool    || '',
        from_programme:   borrowerProgramme || '',
        developer:        developer         || '',
        is_read:          'false',
        created_at:       now,
      };
      await appendRow(SHEET.NOTIFICATIONS, rowToValues('NOTIFICATIONS', notifData));
    }
    res.json({ saved: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ════════════════════════════════════════════════════════════════════════════
//  NOTIFICATIONS
// ════════════════════════════════════════════════════════════════════════════

// NOTE: /read-all must be defined BEFORE /:id/read to avoid route shadowing.
app.patch('/api/notifications/read-all', async (req, res) => {
  try {
    const { school } = req.body || {};
    const rows = await getRows(SHEET.NOTIFICATIONS);
    const toMark = school
      ? rows.filter(r => !r.for_school || r.for_school === school)
      : rows;
    for (const row of toMark) {
      const values = HEADERS['NOTIFICATIONS'].map(h => h === 'is_read' ? 'true' : String(row[h] || ''));
      await updateRow(SHEET.NOTIFICATIONS, row._rowIndex, 'NOTIFICATIONS', values);
    }
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/notifications', async (req, res) => {
  try {
    let rows = await getRows(SHEET.NOTIFICATIONS);
    if (req.query.school) {
      const s = req.query.school;
      rows = rows.filter(r => !r.for_school || r.for_school === s);
    }
    rows.sort((a, b) => new Date(b.created_at || 0) - new Date(a.created_at || 0));
    res.json(rows.slice(0, 50).map(n => ({
      id:            n.id,
      type:          n.type,
      title:         n.title,
      message:       n.message,
      forSchool:     n.for_school,
      forDept:       n.for_dept,
      fromSchool:    n.from_school,
      fromProgramme: n.from_programme,
      developer:     n.developer,
      read:          n.is_read === 'true',
      timestamp:     n.created_at,
    })));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/notifications', async (req, res) => {
  const { type, title, message, forSchool, forDept, fromSchool, fromProgramme, developer } = req.body;
  try {
    const notifData = {
      id:             Date.now(),
      type:           type           || '',
      title:          title          || '',
      message:        message        || '',
      for_school:     forSchool      || '',
      for_dept:       forDept        || '',
      from_school:    fromSchool     || '',
      from_programme: fromProgramme  || '',
      developer:      developer      || '',
      is_read:        'false',
      created_at:     new Date().toISOString(),
    };
    await appendRow(SHEET.NOTIFICATIONS, rowToValues('NOTIFICATIONS', notifData));
    res.json({ saved: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.patch('/api/notifications/:id/read', async (req, res) => {
  try {
    const rows = await getRows(SHEET.NOTIFICATIONS);
    const row  = rows.find(r => String(r.id) === String(req.params.id));
    if (!row) return res.status(404).json({ error: 'not found' });
    const values = HEADERS['NOTIFICATIONS'].map(h => h === 'is_read' ? 'true' : String(row[h] || ''));
    await updateRow(SHEET.NOTIFICATIONS, row._rowIndex, 'NOTIFICATIONS', values);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ════════════════════════════════════════════════════════════════════════════
//  COURSE DETAILS
// ════════════════════════════════════════════════════════════════════════════

// Batch fetch all course details for a programme
app.get('/api/course-details', async (req, res) => {
  const { programme } = req.query;
  try {
    const rows = await getRows(SHEET.DETAILS);
    const filtered = programme
      ? rows.filter(r => r.programme === programme)
      : rows;
    res.json(filtered.map(r => ({
      code:      r.code,
      name:      r.name,
      programme: r.programme,
      school:    r.school,
      dept:      r.dept,
      tier:      parseInt(r.tier) || 4,
      text:      r.content,
      updatedAt: r.updated_at,
    })));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/course-detail', async (req, res) => {
  const { code, name } = req.query;
  try {
    const rows = await getRows(SHEET.DETAILS);
    let row;
    if (code) {
      row = rows
        .filter(r => r.code === code)
        .sort((a, b) => new Date(b.updated_at || 0) - new Date(a.updated_at || 0))[0];
    } else {
      row = rows
        .filter(r => (r.name || '').toLowerCase() === (name || '').toLowerCase())
        .sort((a, b) => new Date(b.updated_at || 0) - new Date(a.updated_at || 0))[0];
    }
    if (!row) return res.status(404).json({ error: 'not found' });
    res.json({
      code:      row.code,
      name:      row.name,
      programme: row.programme,
      school:    row.school,
      dept:      row.dept,
      tier:      parseInt(row.tier) || 4,
      text:      row.content,
      updatedAt: row.updated_at,
    });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/course-detail', async (req, res) => {
  const { code, name, programme, school, dept, tier, text } = req.body;
  if (!code || !text) return res.status(400).json({ error: 'code and text required' });
  try {
    const now  = new Date().toISOString();
    const rows = await getRows(SHEET.DETAILS);
    const existing = rows.find(r => r.code === code && r.programme === (programme || ''));
    const rowData  = {
      id:         existing ? existing.id : Date.now(),
      code,
      name:       name       || '',
      programme:  programme  || '',
      school:     school     || '',
      dept:       dept       || '',
      tier:       tier       || 4,
      content:    text,           // safeValues() will truncate if > 48 000 chars
      updated_at: now,
    };
    const values = rowToValues('DETAILS', rowData);
    if (existing) {
      await updateRow(SHEET.DETAILS, existing._rowIndex, 'DETAILS', values);
    } else {
      await appendRow(SHEET.DETAILS, values);
    }
    res.json({ saved: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ════════════════════════════════════════════════════════════════════════════
//  PROGRAMMES
// ════════════════════════════════════════════════════════════════════════════

/* ── Health check — used by frontend to wake up Render cold-start ── */
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', ts: Date.now(), service: 'kiu-cbe-backend' });
});

app.get('/api/programmes', async (req, res) => {
  try {
    const rows = await getRows(SHEET.PROGRAMMES);
    rows.sort((a, b) => new Date(b.saved_at || 0) - new Date(a.saved_at || 0));
    res.json(rows.map(toProg));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/programmes/:id', async (req, res) => {
  try {
    const rows = await getRows(SHEET.PROGRAMMES);
    const row  = rows.find(r => String(r.id) === String(req.params.id));
    if (!row) return res.status(404).json({ error: 'not found' });
    res.json(toProg(row));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/programmes', async (req, res) => {
  const p = req.body;
  if (!p?.id) return res.status(400).json({ error: 'id required' });
  const meta           = p.meta    || {};
  const courses        = p.courses || [];
  const sectionContent = p.sectionContent || {};
  const now = p.savedAt ? new Date(p.savedAt).toISOString() : new Date().toISOString();
  try {
    const rows     = await getRows(SHEET.PROGRAMMES);
    const existing = rows.find(r => String(r.id) === String(p.id));
    const rowData  = {
      id:               String(p.id),
      name:             meta.name   || p.name   || '',
      abbr:             meta.abbr   || p.abbr   || '',
      school:           meta.school || p.school || '',
      dept:             meta.dept   || p.dept   || '',
      meta:             JSON.stringify(meta),
      courses:          JSON.stringify(courses),
      section_done:     JSON.stringify(p.sectionDone  || {}),
      course_done:      JSON.stringify(p.courseDone   || {}),
      course_count:     p.courseCount || courses.length,
      section_content:  JSON.stringify(sectionContent),
      saved_at:         now,
    };
    const values = rowToValues('PROGRAMMES', rowData);
    if (existing) {
      await updateRow(SHEET.PROGRAMMES, existing._rowIndex, 'PROGRAMMES', values);
    } else {
      await appendRow(SHEET.PROGRAMMES, values);
    }
    res.json({ saved: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ════════════════════════════════════════════════════════════════════════════
//  REGISTRY AUDIT
// ════════════════════════════════════════════════════════════════════════════

app.get('/api/audit', async (req, res) => {
  const filter = req.query.filter;
  try {
    const reg    = await getRows(SHEET.REGISTRY);
    const byName = {};
    reg.forEach(c => {
      const key = (c.name || '').toLowerCase().replace(/[^a-z0-9 ]/g, '').trim();
      if (!key || key.length < 4) return;
      if (!byName[key]) byName[key] = [];
      byName[key].push(toReg(c));
    });
    const conflicts = [];
    Object.entries(byName).forEach(([nameKey, entries]) => {
      if (entries.length < 2) return;
      const codes   = [...new Set(entries.map(e => e.code).filter(Boolean))];
      const nhs     = [...new Set(entries.map(e => (e.DH||0)+(e.SH||0)+(e.AH||0)+(e.OH||0)).filter(v => v > 0))];
      const tiers   = [...new Set(entries.map(e => e.tier || 4))];
      const schools = [...new Set(entries.map(e => e.school).filter(Boolean))];
      const progs   = [...new Set(entries.map(e => e.programme).filter(Boolean))];
      const isCrossSchool   = schools.length > 1;
      const hasCodeConflict = codes.length > 1;
      const hasNHConflict   = nhs.length > 1;
      const hasTierConflict = tiers.length > 1;
      if (filter === 'cross' && !isCrossSchool) return;
      if (!hasCodeConflict && !hasNHConflict && !hasTierConflict && progs.length < 2) return;
      conflicts.push({
        name:           entries[0].name || nameKey,
        nameKey,
        severity:       isCrossSchool ? 'high' : hasCodeConflict || hasNHConflict ? 'medium' : 'low',
        isCrossSchool,
        isCrossDept:    !isCrossSchool && progs.length > 1,
        hasCodeConflict, hasNHConflict, hasTierConflict,
        codes, nhs, tiers, schools, programmes: progs,
        entries: entries.slice(0, 10),
      });
    });
    conflicts.sort((a, b) => {
      const s = { high: 0, medium: 1, low: 2 };
      return (s[a.severity] || 2) - (s[b.severity] || 2);
    });
    res.json({
      total:        conflicts.length,
      high:         conflicts.filter(c => c.severity === 'high').length,
      medium:       conflicts.filter(c => c.severity === 'medium').length,
      low:          conflicts.filter(c => c.severity === 'low').length,
      registrySize: reg.length,
      conflicts,
    });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/audit/resolve', async (req, res) => {
  const { nameKey, canonicalCode, canonicalTier, canonicalDH, canonicalSH, canonicalAH, canonicalOH } = req.body;
  if (!nameKey) return res.status(400).json({ error: 'nameKey required' });
  try {
    const rows     = await getRows(SHEET.REGISTRY);
    const needle   = nameKey.toLowerCase().replace(/[^a-z0-9 ]/g, '').trim();
    let updatedCount = 0;
    for (const row of rows) {
      const rowName = (row.name || '').toLowerCase().replace(/[^a-z0-9 ]/g, '').trim();
      if (rowName !== needle) continue;
      const values = HEADERS['REGISTRY'].map(h => {
        if (h === 'code')       return canonicalCode   || row.code   || '';
        if (h === 'tier')       return String(canonicalTier || 4);
        if (h === 'dh')         return String(canonicalDH   || 0);
        if (h === 'sh')         return String(canonicalSH   || 0);
        if (h === 'ah')         return String(canonicalAH   || 0);
        if (h === 'oh')         return String(canonicalOH   || 0);
        if (h === 'updated_at') return new Date().toISOString();
        return String(row[h] || '');
      });
      await updateRow(SHEET.REGISTRY, row._rowIndex, 'REGISTRY', values);
      updatedCount++;
    }
    const nh = (canonicalDH||0) + (canonicalSH||0) + (canonicalAH||0) + (canonicalOH||0);
    const notifData = {
      id:             Date.now(),
      type:           'AUDIT_RESOLVED',
      title:          `Registry resolved: "${nameKey}"`,
      message:        `Canonical values set: code=${canonicalCode}, ${nh} NH. Update your programme to match.`,
      for_school:     '',  for_dept:       '',
      from_school:    '',  from_programme: '',
      developer:      '',  is_read:        'false',
      created_at:     new Date().toISOString(),
    };
    await appendRow(SHEET.NOTIFICATIONS, rowToValues('NOTIFICATIONS', notifData));
    res.json({ updated: updatedCount });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ════════════════════════════════════════════════════════════════════════════
//  CATALOGUE
// ════════════════════════════════════════════════════════════════════════════

app.get('/api/catalogue', async (req, res) => {
  try {
    const rows   = await getRows(SHEET.REGISTRY);
    const byTier = { 1: 0, 2: 0, 3: 0, 4: 0 };
    rows.forEach(r => {
      const t = parseInt(r.tier) || 4;
      if (byTier[t] !== undefined) byTier[t]++;
    });
    res.json({ byTier, total: rows.length });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ════════════════════════════════════════════════════════════════════════════
//  OWNERSHIP
// ════════════════════════════════════════════════════════════════════════════

app.get('/api/ownership', async (req, res) => {
  try {
    const rows   = await getRows(SHEET.REGISTRY);
    const result = {};
    rows
      .filter(r => parseInt(r.tier) < 4)
      .forEach(r => {
        const key = (r.code || r.name || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        result[key] = {
          code:           r.code,
          name:           r.name,
          ownerSchool:    r.owner_school,
          ownerDept:      r.owner_dept,
          ownerProgramme: r.programme,
          tier:           parseInt(r.tier) || 4,
          developer:      r.developer,
          updatedAt:      r.updated_at,
        };
      });
    res.json(result);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/ownership', async (req, res) => {
  const { courseCode, courseName, ownerSchool, ownerDept, tier } = req.body;
  try {
    const rows    = await getRows(SHEET.REGISTRY);
    const targets = rows.filter(r =>
      (courseCode && r.code.toUpperCase() === courseCode.toUpperCase()) ||
      (!courseCode && (r.name || '').toLowerCase() === (courseName || '').toLowerCase())
    );
    for (const row of targets) {
      const values = HEADERS['REGISTRY'].map(h => {
        if (h === 'tier')         return String(tier || 3);
        if (h === 'owner_school') return ownerSchool || '';
        if (h === 'owner_dept')   return ownerDept   || '';
        if (h === 'updated_at')   return new Date().toISOString();
        return String(row[h] || '');
      });
      await updateRow(SHEET.REGISTRY, row._rowIndex, 'REGISTRY', values);
    }
    res.json({ saved: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Fetch institution website content ────────────────────────────────────────
// ════════════════════════════════════════════════════════════════════════════
//  PROGRAMME SECTIONS  (one row per section — no cell size limit)
// ════════════════════════════════════════════════════════════════════════════

// ── Migrate sections from truncated section_done blob → ProgrammeSections tab ──
app.post('/api/migrate-sections/:programme_id', async (req, res) => {
  const { programme_id } = req.params;
  try {
    // Read raw Programmes rows
    const progRows = await getRows(SHEET.PROGRAMMES);
    const prog = progRows.find(r => String(r.id) === String(programme_id));
    if (!prog) return res.status(404).json({ error: 'Programme not found' });

    const raw = prog.section_done || prog.section_content || '';
    if (!raw) return res.json({ recovered: 0, message: 'No section_done data found' });

    // Remove truncation marker and try full parse first
    const cleaned = raw.replace(/…\[truncated\]/g, '').replace(/\.\.\.\[truncated\]/g, '');
    let sc = {};
    try { sc = JSON.parse(cleaned); } catch(e) {
      // Partial JSON — extract sections individually using string search
      const keys = ['intro','governance','nameCode','description','development',
        'rationale','competences','outcomes','targetGroup','admission',
        'humanRes','infrastructure','delivery','assessment','progression',
        'gradLoad','research','community','policies','financial','welfare'];
      for (const key of keys) {
        const startMarker = `"${key}":{"text":"`;
        const startIdx = cleaned.indexOf(startMarker);
        if (startIdx < 0) continue;
        const textStart = startIdx + startMarker.length;
        // Find end of text value — look for ","done":
        let depth = 0, inStr = false, escape = false;
        let textEnd = -1;
        for (let i = textStart; i < cleaned.length; i++) {
          const ch = cleaned[i];
          if (escape) { escape = false; continue; }
          if (ch === '\\') { escape = true; continue; }
          if (ch === '"' && !inStr) { inStr = true; continue; }
          if (ch === '"' && inStr) {
            // Check if followed by ,"done":
            const ahead = cleaned.slice(i+1, i+10);
            if (ahead.startsWith('","done"') || ahead.startsWith('"')) {
              textEnd = i; break;
            }
          }
        }
        if (textEnd > textStart) {
          const text = cleaned.slice(textStart, textEnd)
            .replace(/\\n/g, '\n').replace(/\\"/g, '"').replace(/\\\\/g, '\\');
          if (text.length > 20) {
            sc[key] = { text, done: true };
          }
        }
      }
    }

    if (!Object.keys(sc).length) {
      return res.json({ recovered: 0, message: 'Could not extract any sections from backup' });
    }

    // Save recovered sections to ProgrammeSections tab
    const now = new Date().toISOString();
    let saved = 0;
    const existingRows = await getRows(SHEET.SECTIONS);
    for (const [key, sec] of Object.entries(sc)) {
      if (!sec.text || sec.text.length < 20) continue;
      const existing = existingRows.find(r =>
        String(r.programme_id) === String(programme_id) && r.section_key === key
      );
      const rowData = {
        id:             `${programme_id}_${key}`,
        programme_id:   String(programme_id),
        programme_name: prog.name || '',
        section_key:    key,
        content:        sec.text,
        done:           'true',
        updated_at:     now,
      };
      const values = rowToValues('SECTIONS', rowData);
      if (existing) {
        await updateRow(SHEET.SECTIONS, existing._rowIndex, 'SECTIONS', values);
      } else {
        await appendRow(SHEET.SECTIONS, values);
      }
      saved++;
    }
    res.json({ recovered: saved, sections: Object.keys(sc) });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/sections', async (req, res) => {
  const { programme_id } = req.query;
  try {
    const rows = await getRows(SHEET.SECTIONS);
    const filtered = programme_id
      ? rows.filter(r => String(r.programme_id) === String(programme_id))
      : rows;
    const sc = {};
    filtered.forEach(r => {
      if (r.section_key) {
        sc[r.section_key] = { text: r.content || '', done: r.done === 'true' };
      }
    });
    res.json(sc);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Batch save all sections for a programme (1 read instead of 21) ───────────
app.post('/api/sections/batch', async (req, res) => {
  const { programme_id, programme_name, sections } = req.body;
  if (!programme_id || !Array.isArray(sections))
    return res.status(400).json({ error: 'programme_id and sections array required' });
  try {
    const now  = new Date().toISOString();
    // ONE read to get all existing section rows for this programme
    const rows = await getRows(SHEET.SECTIONS);
    const existingMap = {};
    rows.filter(r => String(r.programme_id) === String(programme_id))
        .forEach(r => { existingMap[r.section_key] = r; });

    let saved = 0;
    for (const sec of sections) {
      if (!sec.section_key || !sec.content) continue;
      const rowData = {
        id:             `${programme_id}_${sec.section_key}`,
        programme_id:   String(programme_id),
        programme_name: programme_name || '',
        section_key:    sec.section_key,
        content:        sec.content,
        done:           String(!!sec.done),
        updated_at:     now,
      };
      const values = rowToValues('SECTIONS', rowData);
      if (existingMap[sec.section_key]) {
        await updateRow(SHEET.SECTIONS, existingMap[sec.section_key]._rowIndex, 'SECTIONS', values);
      } else {
        await appendRow(SHEET.SECTIONS, values);
      }
      saved++;
    }
    res.json({ saved });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Batch save all course details for a programme (1 read instead of 30) ─────
app.post('/api/course-details/batch', async (req, res) => {
  const { programme, items } = req.body;
  if (!programme || !Array.isArray(items))
    return res.status(400).json({ error: 'programme and items array required' });
  try {
    const now  = new Date().toISOString();
    // ONE read to get all existing course detail rows for this programme
    const rows = await getRows(SHEET.DETAILS);
    const existingMap = {};
    rows.filter(r => r.programme === programme)
        .forEach(r => { existingMap[r.code] = r; });

    let saved = 0;
    for (const item of items) {
      if (!item.code || !item.text) continue;
      const rowData = {
        id:         existingMap[item.code] ? existingMap[item.code].id : `${Date.now()}_${saved}`,
        code:       item.code,
        name:       item.name || '',
        programme:  item.programme || programme,
        school:     item.school || '',
        dept:       item.dept   || '',
        tier:       item.tier   || 4,
        content:    item.text,
        updated_at: now,
      };
      const values = rowToValues('DETAILS', rowData);
      if (existingMap[item.code]) {
        await updateRow(SHEET.DETAILS, existingMap[item.code]._rowIndex, 'DETAILS', values);
      } else {
        await appendRow(SHEET.DETAILS, values);
      }
      saved++;
    }
    res.json({ saved });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/section', async (req, res) => {
  const { programme_id, programme_name, section_key, content, done } = req.body;
  if (!programme_id || !section_key) return res.status(400).json({ error: 'programme_id and section_key required' });
  try {
    const now = new Date().toISOString();
    const rows = await getRows(SHEET.SECTIONS);
    const existing = rows.find(r =>
      String(r.programme_id) === String(programme_id) && r.section_key === section_key
    );
    const rowData = {
      id:             existing ? existing.id : `${programme_id}_${section_key}`,
      programme_id:   String(programme_id),
      programme_name: programme_name || '',
      section_key:    section_key,
      content:        content || '',
      done:           String(!!done),
      updated_at:     now,
    };
    const values = rowToValues('SECTIONS', rowData);
    if (existing) {
      await updateRow(SHEET.SECTIONS, existing._rowIndex, 'SECTIONS', values);
    } else {
      await appendRow(SHEET.SECTIONS, values);
    }
    res.json({ saved: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Delete a programme row from Google Sheets
app.delete('/api/programmes/:id', async (req, res) => {
  const { id } = req.params;
  try {
    const rows = await getRows(SHEET.PROGRAMMES);
    const row  = rows.find(r => String(r.id) === String(id));
    if (!row) return res.status(404).json({ error: 'not found' });
    // Delete by clearing the row (Google Sheets API doesn't have native delete row without batchUpdate)
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{
          deleteDimension: {
            range: {
              sheetId: await getSheetId(SHEET.PROGRAMMES),
              dimension: 'ROWS',
              startIndex: row._rowIndex - 1,  // 0-based
              endIndex:   row._rowIndex,
            }
          }
        }]
      }
    });
    res.json({ deleted: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Delete all sections for a programme from ProgrammeSections tab
app.delete('/api/sections/:programme_id', async (req, res) => {
  const { programme_id } = req.params;
  try {
    const rows = await getRows(SHEET.SECTIONS);
    const toDelete = rows
      .filter(r => String(r.programme_id) === String(programme_id))
      .sort((a, b) => b._rowIndex - a._rowIndex); // delete from bottom up
    const sheetId = await getSheetId(SHEET.SECTIONS);
    for (const row of toDelete) {
      await sheetsClient.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          requests: [{
            deleteDimension: {
              range: {
                sheetId,
                dimension: 'ROWS',
                startIndex: row._rowIndex - 1,
                endIndex:   row._rowIndex,
              }
            }
          }]
        }
      });
    }
    res.json({ deleted: toDelete.length });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Helper — get the internal sheetId (gid) for a named tab
async function getSheetId(tabName) {
  const meta = await sheetsClient.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const sheet = meta.data.sheets.find(s => s.properties.title === tabName);
  if (!sheet) throw new Error(`Tab "${tabName}" not found`);
  return sheet.properties.sheetId;
}

app.get('/api/fetch-institution', async (req, res) => {
  const { url } = req.query;
  if (!url) return res.status(400).json({ error: 'url required' });
  try {
    const r = await fetch(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; KIU-CBE-Bot/1.0)',
        'Accept': 'text/html,application/xhtml+xml',
      },
      signal: AbortSignal.timeout(12000),
    });
    if (!r.ok) return res.status(r.status).json({ error: 'Fetch failed: '+r.status });
    const html = await r.text();
    // Strip HTML tags and clean up whitespace
    const text = html
      .replace(/<script[\s\S]*?<\/script>/gi, ' ')
      .replace(/<style[\s\S]*?<\/style>/gi, ' ')
      .replace(/<[^>]+>/g, ' ')
      .replace(/\s+/g, ' ')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .trim()
      .slice(0, 6000);
    res.json({ content: text, url, length: text.length });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── Retry wrapper for initSheets — handles Render cold-start network hiccups ──
async function initSheetsWithRetry(maxAttempts = 5, delayMs = 10000) {
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const ok = await initSheets();
    if (ok) return true;
    if (attempt < maxAttempts) {
      console.log(`⏳ Sheets init attempt ${attempt}/${maxAttempts} failed — retrying in ${delayMs/1000}s...`);
      await new Promise(res => setTimeout(res, delayMs));
    }
  }
  console.error(`❌ Sheets init failed after ${maxAttempts} attempts — API will run without storage until next deploy.`);
  return false;
}

// ── Start ─────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`KIU CBE API v4.0 (Google Sheets) — port ${PORT}`);
  console.log(`SPREADSHEET_ID:              ${process.env.SPREADSHEET_ID              ? 'SET ✅' : 'NOT SET ⚠️'}`);
  console.log(`GOOGLE_SERVICE_ACCOUNT_JSON: ${process.env.GOOGLE_SERVICE_ACCOUNT_JSON ? 'SET ✅' : 'NOT SET ⚠️'}`);
  console.log(`OPENROUTER_API_KEY:          ${process.env.OPENROUTER_API_KEY          ? 'SET ✅' : 'NOT SET ⚠️'}`);
  initSheetsWithRetry(5, 10000);
});
