require('dotenv').config();
const express = require('express');
const session = require('express-session');
const multer = require('multer');
const { parse } = require('csv-parse/sync');
const { v4: uuidv4 } = require('uuid');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const { S3Client, GetObjectCommand, PutObjectCommand } = require('@aws-sdk/client-s3');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(session({
  secret: process.env.SESSION_SECRET || 'secret',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 8 * 60 * 60 * 1000 },
}));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

// ── TikTok config ──────────────────────────────────────────────────────────
const TT_BASE = 'https://business-api.tiktok.com/open_api/v1.3';
const ADV_ID = process.env.TIKTOK_ADVERTISER_ID;
const PIXEL_ID = process.env.TIKTOK_PIXEL_ID;
const BC_ID = process.env.TIKTOK_BC_ID;
const APP_ID = process.env.TIKTOK_APP_ID;
const APP_SECRET = process.env.TIKTOK_APP_SECRET;
const REDIRECT_URI = process.env.TIKTOK_REDIRECT_URI || 'https://campaigns.videosapi.net/auth/callback';

let activeToken = process.env.TIKTOK_ACCESS_TOKEN || null;
function getToken() { return activeToken; }

// adv_id defaults to env ADV_ID so existing calls work unchanged
async function ttGet(path, params = {}, adv_id = ADV_ID) {
  const url = new URL(`${TT_BASE}${path}`);
  if (adv_id) url.searchParams.set('advertiser_id', adv_id);
  Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, typeof v === 'object' ? JSON.stringify(v) : v));
  const res = await fetch(url.toString(), { headers: { 'Access-Token': getToken() } });
  return res.json();
}

async function ttPost(path, body, adv_id = ADV_ID) {
  const res = await fetch(`${TT_BASE}${path}`, {
    method: 'POST',
    headers: { 'Access-Token': getToken(), 'Content-Type': 'application/json' },
    body: JSON.stringify({ advertiser_id: adv_id, ...body }),
  });
  return res.json();
}

// ── TikTok OAuth ───────────────────────────────────────────────────────────
app.get('/auth', requireAuth, (req, res) => {
  const state = uuidv4();
  req.session.oauthState = state;
  const url = new URL('https://business-api.tiktok.com/portal/auth');
  url.searchParams.set('app_id', APP_ID);
  url.searchParams.set('state', state);
  url.searchParams.set('redirect_uri', REDIRECT_URI);
  res.redirect(url.toString());
});

app.post('/api/exchange-token', requireAuth, async (req, res) => {
  const { auth_code } = req.body;
  if (!auth_code) return res.status(400).json({ error: 'auth_code required' });
  try {
    const r = await fetch(`${TT_BASE}/oauth2/access_token/`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ app_id: APP_ID, secret: APP_SECRET, auth_code }),
    });
    const data = await r.json();
    if (data.data?.access_token) {
      activeToken = data.data.access_token;
      pixelCache = {}; // clear cache on token refresh
      console.log('TikTok token exchanged successfully');
      res.json({ ok: true });
    } else {
      res.status(400).json({ error: `TikTok error ${data.code}: ${data.message}` });
    }
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/token-status', requireAuth, async (req, res) => {
  if (!getToken()) return res.json({ valid: false, reason: 'No token set' });
  try {
    const data = await ttGet('/identity/get/');
    res.json({ valid: data.code === 0, code: data.code, message: data.message });
  } catch (e) {
    res.json({ valid: false, reason: e.message });
  }
});

// GeoNames ID map for TikTok location targeting
const GEO_MAP = {
  US: '6252001', GB: '2635167', UK: '2635167', AU: '2077456', CA: '6251999',
  DE: '2921044', FR: '3017382', MX: '3996063', BR: '3469034', IN: '1269750',
  JP: '1861060', SG: '1880251', PH: '1694008', ID: '1643084', TH: '1605651',
  MY: '1733045', VN: '1562822', ZA: '953987', NG: '2328926', KE: '192950',
  AE: '290557', SA: '102358', NZ: '2186224', IT: '3175395', ES: '2510769',
  NL: '2750405', SE: '2661886', NO: '607072', DK: '2623032', FI: '660013',
  PL: '798544', TR: '298795', AR: '3865483', CO: '3686110', CL: '3895114',
  PE: '3932488', EG: '357994', MA: '2542007', GH: '2300660', TZ: '149590',
  UA: '690791', RO: '798549', HU: '719819', CZ: '3077311', GR: '390903',
  PT: '2264397', BE: '2802361', AT: '2782113', CH: '2658434', IL: '294640',
  KR: '1835841', TW: '1668284', HK: '1819730', PK: '1168579', BD: '1210997',
};

// ── Auth helpers ───────────────────────────────────────────────────────────
function parseUsers() {
  const map = {};
  (process.env.USERS || '').split(',').forEach(entry => {
    const [u, ...rest] = entry.trim().split(':');
    if (u) map[u.trim()] = rest.join(':').trim();
  });
  return map;
}

function requireAuth(req, res, next) {
  if (req.session.user) return next();
  res.status(401).json({ error: 'Unauthorized' });
}

// ── Persistent log store (Cloudflare R2) ──────────────────────────────────
const R2_BUCKET = process.env.R2_BUCKET_NAME;
const LOG_KEY   = 'campaign-logs.json';

const r2 = R2_BUCKET ? new S3Client({
  region: 'auto',
  endpoint: `https://${process.env.R2_ACCOUNT_ID}.r2.cloudflarestorage.com`,
  credentials: {
    accessKeyId:     process.env.R2_ACCESS_KEY_ID,
    secretAccessKey: process.env.R2_SECRET_ACCESS_KEY,
  },
}) : null;

async function loadLogs() {
  if (r2) {
    try {
      const res = await r2.send(new GetObjectCommand({ Bucket: R2_BUCKET, Key: LOG_KEY }));
      const body = await res.Body.transformToString();
      return JSON.parse(body);
    } catch (e) {
      if (e.name !== 'NoSuchKey') console.error('R2 loadLogs error:', e.message);
      return [];
    }
  }
  // Fallback: local file
  const LOCAL = process.env.LOG_PATH || path.join(__dirname, 'campaign-logs.json');
  try { if (fs.existsSync(LOCAL)) return JSON.parse(fs.readFileSync(LOCAL, 'utf8')); } catch (_) {}
  return [];
}

async function appendLog(entry) {
  const logs = await loadLogs();
  logs.push(entry);
  const body = JSON.stringify(logs, null, 2);
  if (r2) {
    try {
      await r2.send(new PutObjectCommand({ Bucket: R2_BUCKET, Key: LOG_KEY, Body: body, ContentType: 'application/json' }));
    } catch (e) { console.error('R2 appendLog error:', e.message); }
  } else {
    const LOCAL = process.env.LOG_PATH || path.join(__dirname, 'campaign-logs.json');
    try { fs.writeFileSync(LOCAL, body); } catch (e) { console.error('Failed to write log:', e.message); }
  }
}

// ── In-memory stores ───────────────────────────────────────────────────────
const jobs = new Map();
let pixelCache = {};     // adv_id -> pixel_id

function jobEmit(jobId, event) {
  const job = jobs.get(jobId);
  if (!job) return;
  job.events.push(event);
  job.clients.forEach(res => res.write(`data: ${JSON.stringify(event)}\n\n`));
}

// ── Routes ─────────────────────────────────────────────────────────────────

app.post('/api/login', (req, res) => {
  const { username, password } = req.body;
  const users = parseUsers();
  if (users[username] && users[username] === password) {
    req.session.user = username;
    res.json({ ok: true });
  } else {
    res.status(401).json({ error: 'Invalid credentials' });
  }
});

app.post('/api/logout', (req, res) => { req.session.destroy(); res.json({ ok: true }); });
app.get('/api/me', (req, res) => { res.json({ user: req.session.user || null }); });

app.get('/api/identities', requireAuth, async (req, res) => {
  try {
    const data = await ttGet('/identity/get/');
    if (data.code !== 0) return res.status(502).json({ error: `TikTok API error ${data.code}: ${data.message}` });
    const list = (data.data?.identity_list || []).map(i => ({
      id: i.identity_id,
      name: i.display_name,
      type: i.identity_type,
      bc_id: i.identity_authorized_bc_id,
      avatar: i.profile_image,
    }));
    res.json(list);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/accounts', requireAuth, async (req, res) => {
  try {
    const url = new URL(`${TT_BASE}/oauth2/advertiser/get/`);
    url.searchParams.set('app_id', APP_ID);
    url.searchParams.set('secret', APP_SECRET);
    const r = await fetch(url.toString(), { headers: { 'Access-Token': getToken() } });
    const data = await r.json();
    if (data.code !== 0) return res.status(502).json({ error: `TikTok API error ${data.code}: ${data.message}` });
    const list = (data.data?.list || []).map(a => ({
      id: String(a.advertiser_id),
      name: a.advertiser_name,
    }));
    // Always include the env default account if not already in list
    if (ADV_ID && !list.find(a => a.id === ADV_ID)) {
      try {
        const info = await ttGet('/advertiser/info/', { fields: JSON.stringify(['advertiser_name']) });
        const name = info.data?.list?.[0]?.advertiser_name || `Account ${ADV_ID}`;
        list.unshift({ id: ADV_ID, name });
      } catch (_) {
        list.unshift({ id: ADV_ID, name: `Account ${ADV_ID}` });
      }
    }
    res.json(list);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/logs', requireAuth, async (req, res) => {
  try {
    res.json((await loadLogs()).reverse());
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/sample', requireAuth, async (req, res) => {
  // Fetch live account and identity names for dropdowns
  let accountNames = [];
  let identityNames = [];

  try {
    const url = new URL(`${TT_BASE}/oauth2/advertiser/get/`);
    url.searchParams.set('app_id', APP_ID);
    url.searchParams.set('secret', APP_SECRET);
    const r = await fetch(url.toString(), { headers: { 'Access-Token': getToken() } });
    const data = await r.json();
    accountNames = (data.data?.list || []).map(a => a.advertiser_name);
    if (!accountNames.length && ADV_ID) {
      const info = await ttGet('/advertiser/info/', { fields: JSON.stringify(['advertiser_name']) });
      const name = info.data?.list?.[0]?.advertiser_name;
      if (name) accountNames = [name];
    }
  } catch (_) {}

  try {
    const data = await ttGet('/identity/get/');
    identityNames = (data.data?.identity_list || []).map(i => i.display_name);
  } catch (_) {}

  if (!accountNames.length) accountNames = ['My Account Name'];
  if (!identityNames.length) identityNames = ['My TikTok Identity'];

  const geos = Object.keys(GEO_MAP);

  // Language names from TikTok /tool/language/ — keep in sync with LANGUAGE_NAME_TO_CODE
  const LANGUAGE_NAMES = [
    'Arabic','Assamese','Haryanvi','Bihari','Bengali','Czech','German','Greek',
    'English','Spanish','Finnish','French','Gujarati','Hebrew','Hindi','Hungarian',
    'Indonesian','Italian','Japanese','Khmer','Kannada','Korean','Malayalam',
    'Marathi','Malay','Dutch','Oriya','Punjabi','Polish','Portuguese','Rajasthani',
    'Romanian','Russian','Swedish','Tamil','Telugu','Thai','Turkish','Ukrainian',
    'Vietnamese','Simplified Chinese','Traditional Chinese',
  ];

  const wb = new ExcelJS.Workbook();

  // ── Hidden lookups sheet (source for dropdown ranges) ──
  const lookups = wb.addWorksheet('_Lookups');
  lookups.state = 'veryHidden';
  accountNames.forEach((n, i) => { lookups.getCell(i + 1, 1).value = n; });
  identityNames.forEach((n, i) => { lookups.getCell(i + 1, 2).value = n; });
  geos.forEach((g, i) => { lookups.getCell(i + 1, 3).value = g; });
  LANGUAGE_NAMES.forEach((n, i) => { lookups.getCell(i + 1, 4).value = n; });

  // ── Main campaigns sheet ──
  const ws = wb.addWorksheet('Campaigns');

  const HEADERS = [
    { key: 'account_name',  label: 'account_name',  width: 28 },
    { key: 'identity_name', label: 'identity_name',  width: 28 },
    { key: 'campaign_name', label: 'campaign_name',  width: 30 },
    { key: 'geo',           label: 'geo',            width: 10 },
    { key: 'budget',        label: 'budget',         width: 10 },
    { key: 'bid_strategy',  label: 'bid_strategy',   width: 16 },
    { key: 'bid_amount',    label: 'bid_amount',     width: 12 },
    { key: 'targeting',     label: 'targeting',      width: 16 },
    { key: 'language',      label: 'language',       width: 18 },
    { key: 'start_date',    label: 'start_date',     width: 14 },
    { key: 'start_time',    label: 'start_time',     width: 12 },
    ...Array.from({length:10},(_,i) => ({ key:`video_url_${i+1}`, label:`video_url_${i+1}`, width: 44 })),
    ...Array.from({length:5}, (_,i) => ({ key:`headline_${i+1}`,  label:`headline_${i+1}`,  width: 32 })),
    { key: 'url', label: 'url', width: 50 },
  ];

  ws.columns = HEADERS.map(h => ({ header: h.label, key: h.key, width: h.width }));

  // Style header row
  ws.getRow(1).eachCell(cell => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2d2d3f' } };
    cell.alignment = { vertical: 'middle' };
  });
  ws.getRow(1).height = 20;

  // Build colIndex first so it's available for both numFmt and data validation
  const DATA_ROWS = 500;
  const colIndex = {};
  HEADERS.forEach((h, i) => { colIndex[h.key] = i + 1; });

  // Set start_date column as date format so Excel shows a calendar picker
  ws.getColumn(colIndex.start_date).numFmt = 'dd/mm/yyyy';

  const exampleDate = new Date(2026, 6, 10); // July 10 2026 (month is 0-indexed)
  ws.addRow({
    account_name: accountNames[0], identity_name: identityNames[0],
    campaign_name: 'Example_Campaign_Jul', geo: 'US', budget: 30,
    bid_strategy: 'LOWEST_COST', bid_amount: '', targeting: 'BROAD', language: 'English',
    start_date: exampleDate, start_time: '00:00',
    video_url_1: 'https://videosapi.net/videos/example1.mp4',
    headline_1: 'Your headline here', headline_2: 'Second headline',
    url: 'https://yoursite.com/landing',
  });
  ws.addRow({
    account_name: accountNames[0], identity_name: identityNames[0],
    campaign_name: 'Example_Campaign_2_Jul', geo: 'US', budget: 50,
    bid_strategy: 'COST_CAP', bid_amount: 0.75, targeting: 'AGE_35_PLUS', language: 'German',
    start_date: exampleDate, start_time: '08:00',
    video_url_1: 'https://videosapi.net/videos/example2.mp4',
    headline_1: 'Another headline', headline_2: 'Try it today',
    url: 'https://yoursite.com/offer',
  });

  // ── Data validation dropdowns for rows 2–500 ──

  const acctRef   = `'_Lookups'!$A$1:$A$${accountNames.length}`;
  const identRef  = `'_Lookups'!$B$1:$B$${identityNames.length}`;
  const geoRef    = `'_Lookups'!$C$1:$C$${geos.length}`;
  const langRef   = `'_Lookups'!$D$1:$D$${LANGUAGE_NAMES.length}`;

  for (let row = 2; row <= DATA_ROWS + 1; row++) {
    ws.getCell(row, colIndex.account_name).dataValidation  = { type: 'list', allowBlank: true, formulae: [acctRef] };
    ws.getCell(row, colIndex.identity_name).dataValidation = { type: 'list', allowBlank: true, formulae: [identRef] };
    ws.getCell(row, colIndex.geo).dataValidation           = { type: 'list', allowBlank: true, formulae: [geoRef] };
    ws.getCell(row, colIndex.bid_strategy).dataValidation  = { type: 'list', allowBlank: true, formulae: ['"LOWEST_COST,COST_CAP"'] };
    ws.getCell(row, colIndex.targeting).dataValidation     = { type: 'list', allowBlank: true, formulae: ['"BROAD,AGE_35_PLUS"'] };
    ws.getCell(row, colIndex.language).dataValidation      = { type: 'list', allowBlank: true, formulae: [langRef] };
    // Date validation so blank cells also inherit calendar format
    ws.getCell(row, colIndex.start_date).dataValidation    = { type: 'date', allowBlank: true, operator: 'greaterThan', formulae: [new Date(2020, 0, 1)] };
    ws.getCell(row, colIndex.start_date).numFmt = 'dd/mm/yyyy';
  }

  // Freeze header row
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  res.setHeader('Content-Disposition', 'attachment; filename="campaign-template.xlsx"');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await wb.xlsx.write(res);
  res.end();
});

app.post('/api/parse-csv', requireAuth, upload.single('csv'), async (req, res) => {
  try {
    let records;
    const isXlsx = req.file.originalname.endsWith('.xlsx') || req.file.originalname.endsWith('.xls');

    if (isXlsx) {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(req.file.buffer);
      const ws = wb.worksheets.find(s => !s.name.startsWith('_')) || wb.worksheets[0];
      const headers = [];
      ws.getRow(1).eachCell((cell, col) => { headers[col] = String(cell.value || '').trim(); });
      records = [];
      ws.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        const obj = {};
        let hasValue = false;
        row.eachCell((cell, col) => {
          const key = headers[col];
          if (!key) return;
          let val;
          const cv = cell.value;
          if (cv instanceof Date) {
            // Excel date cell → YYYY-MM-DD; treat epoch-zero (empty date cell) as blank
            if (cv.getFullYear() <= 1900) {
              val = '';
            } else {
              val = `${cv.getFullYear()}-${String(cv.getMonth()+1).padStart(2,'0')}-${String(cv.getDate()).padStart(2,'0')}`;
            }
          } else if (cv !== null && cv !== undefined && typeof cv === 'object') {
            // Hyperlink cell: { text, hyperlink } or { richText: [{text},...] }
            if (cv.hyperlink) val = cv.hyperlink.trim();
            else if (cv.text) val = String(cv.text).trim();
            else if (Array.isArray(cv.richText)) val = cv.richText.map(r => r.text || '').join('').trim();
            else val = '';
          } else {
            val = cv === null || cv === undefined ? '' : String(cv).trim();
          }
          obj[key] = val;
          if (val) hasValue = true;
        });
        if (hasValue) records.push(obj);
      });
    } else {
      const text = req.file.buffer.toString('utf8');
      records = parse(text, { columns: true, skip_empty_lines: true, trim: true });
    }

    const rows = records.map((r, i) => {
      const videos = [];
      for (let n = 1; n <= 10; n++) {
        const v = r[`video_url_${n}`];
        if (v) videos.push(v);
      }
      const headlines = [];
      for (let n = 1; n <= 5; n++) {
        const h = r[`headline_${n}`];
        if (h) headlines.push(h);
      }
      const geo = (r.geo || '').toUpperCase();
      const bid_strategy = (r.bid_strategy || '').toUpperCase();
      const targeting = (r.targeting || '').toUpperCase();
      const language = (r.language || '').trim(); // language name e.g. "German"
      const budget = parseFloat(r.budget);
      const bid_amount = parseFloat(r.bid_amount) || null;

      const validationErrors = [];
      if (!r.account_name?.trim())          validationErrors.push('account_name is required');
      if (!r.identity_name?.trim())         validationErrors.push('identity_name is required');
      if (!r.campaign_name?.trim())         validationErrors.push('campaign_name is required');
      if (!geo)                             validationErrors.push('geo is required');
      else if (!GEO_MAP[geo])               validationErrors.push(`geo "${geo}" is not a supported country code`);
      if (isNaN(budget) || budget <= 0)     validationErrors.push('budget must be a positive number');
      if (!bid_strategy)                    validationErrors.push('bid_strategy is required');
      else if (!['LOWEST_COST','COST_CAP'].includes(bid_strategy)) validationErrors.push('bid_strategy must be LOWEST_COST or COST_CAP');
      if (bid_strategy === 'COST_CAP' && !bid_amount) validationErrors.push('bid_amount is required when bid_strategy is COST_CAP');
      if (!targeting)                       validationErrors.push('targeting is required');
      else if (!['BROAD','AGE_35_PLUS'].includes(targeting)) validationErrors.push('targeting must be BROAD or AGE_35_PLUS');
      if (!r.start_date?.trim())            validationErrors.push('start_date is required');
      if (!videos.length)                   validationErrors.push('at least one video_url is required');
      if (!headlines.length)                validationErrors.push('at least one headline is required');
      if (!r.url?.trim())                   validationErrors.push('landing page url is required');
      else if (!/^https?:\/\/.+/.test(r.url.trim())) validationErrors.push('url must start with http:// or https://');

      return {
        rowIndex: i,
        account_name: r.account_name || '',
        identity_name: r.identity_name || '',
        campaign_name: r.campaign_name,
        geo,
        budget,
        bid_strategy,
        bid_amount,
        targeting,
        start_date: r.start_date,
        start_time: r.start_time || '00:00',
        videos,
        headlines,
        url: r.url,
        language,
        validationErrors,
        status: 'pending',
        error: null,
        campaign_id: null,
      };
    });

    res.json({ rows });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.post('/api/deploy', requireAuth, async (req, res) => {
  const { rows, accountsMap } = req.body;
  if (!rows?.length) return res.status(400).json({ error: 'No rows' });

  // Build identity map: name -> { id, type, bc_id }
  let identityMap = {};
  try {
    const data = await ttGet('/identity/get/');
    (data.data?.identity_list || []).forEach(i => {
      identityMap[i.display_name] = {
        id: i.identity_id,
        type: i.identity_type,
        bc_id: i.identity_authorized_bc_id,
      };
    });
  } catch (_) {}

  const jobId = uuidv4();
  const deployedBy = req.session.user || 'unknown';
  jobs.set(jobId, { events: [], clients: new Set(), status: 'running' });
  res.json({ jobId });

  deployRows(jobId, rows, accountsMap || {}, identityMap, deployedBy).catch(console.error);
});

app.get('/api/deploy/:jobId/events', requireAuth, (req, res) => {
  const job = jobs.get(req.params.jobId);
  if (!job) return res.status(404).end();

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  job.events.forEach(e => res.write(`data: ${JSON.stringify(e)}\n\n`));
  if (job.status === 'done') return res.end();

  job.clients.add(res);
  req.on('close', () => job.clients.delete(res));
});

// ── Pixel lookup per account ───────────────────────────────────────────────
async function getPixelForAccount(adv_id) {
  if (pixelCache[adv_id]) return pixelCache[adv_id];
  try {
    const data = await ttGet('/pixel/list/get/', {}, adv_id);
    const pixel = (data.data?.list || []).find(p => p.status === 'ACTIVE') || data.data?.list?.[0];
    const pid = pixel?.pixel_id || PIXEL_ID;
    pixelCache[adv_id] = pid;
    return pid;
  } catch (_) {
    return PIXEL_ID;
  }
}

// ── Deployment logic ───────────────────────────────────────────────────────
async function deployRows(jobId, rows, accountsMap, identityMap, deployedBy = 'unknown') {
  for (const row of rows) {
    const adv_id = accountsMap[row.account_name] || ADV_ID;
    const account_name = row.account_name || 'Default';

    // Resolve identity from row's identity_name field
    const resolvedIdentity = identityMap[row.identity_name] || Object.values(identityMap)[0] || {};
    const identity_id = resolvedIdentity.id;
    const identity_type = resolvedIdentity.type || 'BC_AUTH_TT';
    const identity_bc_id = resolvedIdentity.bc_id || BC_ID;

    jobEmit(jobId, { type: 'row_start', rowIndex: row.rowIndex, campaign_name: row.campaign_name });

    const logEntry = {
      timestamp: new Date().toISOString(),
      deployed_by: deployedBy,
      account_name,
      advertiser_id: adv_id,
      campaign_name: row.campaign_name,
      geo: row.geo,
      budget: row.budget,
      bid_strategy: row.bid_strategy,
      bid_amount: row.bid_amount,
      video_count: row.videos.length,
      headline_count: row.headlines.length,
      status: 'error',
      campaign_id: null,
      error: null,
    };

    try {
      const pixel_id = await getPixelForAccount(adv_id);

      jobEmit(jobId, { type: 'step', rowIndex: row.rowIndex, step: 'Creating campaign…' });
      const campaign = await createCampaign(row, adv_id);
      const campaign_id = campaign.data?.campaign_id;
      if (!campaign_id) {
        const msg = campaign.message || JSON.stringify(campaign);
        throw new Error(campaign.code === 40911 || (msg && msg.includes('already'))
          ? `Campaign name "${row.campaign_name}" already exists — rename it in the CSV and retry`
          : `Campaign failed: ${msg}`);
      }

      jobEmit(jobId, { type: 'step', rowIndex: row.rowIndex, step: `Uploading ${row.videos.length} video(s)…` });
      const { ids: video_ids, coverPromises } = await uploadVideos(row.videos, adv_id, jobId, row.rowIndex);

      jobEmit(jobId, { type: 'step', rowIndex: row.rowIndex, step: 'Creating ad group…' });
      const adgroup = await createAdGroup(row, campaign_id, adv_id, pixel_id);
      const adgroup_id = adgroup.data?.adgroup_id;
      if (!adgroup_id) throw new Error(`Ad group failed: ${JSON.stringify(adgroup)}`);

      jobEmit(jobId, { type: 'step', rowIndex: row.rowIndex, step: 'Creating ads…' });
      await createAds(row, adgroup_id, video_ids, identity_id, identity_type, identity_bc_id, coverPromises, adv_id);

      logEntry.status = 'success';
      logEntry.campaign_id = campaign_id;
      jobEmit(jobId, { type: 'row_done', rowIndex: row.rowIndex, campaign_id });

    } catch (err) {
      logEntry.error = err.message;
      jobEmit(jobId, { type: 'row_error', rowIndex: row.rowIndex, error: err.message });
    }

    await appendLog(logEntry);
  }

  jobEmit(jobId, { type: 'done' });
  const job = jobs.get(jobId);
  if (job) {
    job.status = 'done';
    job.clients.forEach(r => r.end());
  }
}

async function createCampaign(row, adv_id) {
  return ttPost('/smart_plus/campaign/create/', {
    request_id: Date.now().toString(),
    campaign_name: row.campaign_name,
    objective_type: 'WEB_CONVERSIONS',
    sales_destination: 'WEBSITE',
    budget_optimize_on: true,
    budget_mode: 'BUDGET_MODE_DYNAMIC_DAILY_BUDGET',
    budget: row.budget,
    operation_status: 'ENABLE',
  }, adv_id);
}

async function findExistingVideo(baseName, adv_id) {
  // Use full baseName (not truncated) so similar filenames don't cross-match
  const baseCore = baseName.replace(/^\d+_/, '').replace(/[_\s-]+/g, ' ').toLowerCase();
  for (let page = 1; page <= 15; page++) {
    try {
      const res = await ttGet('/file/video/ad/search/', { page, page_size: 20 }, adv_id);
      const list = res.data?.list || [];
      const match = list.find(v => {
        const norm = (v.file_name || '').replace(/[_\s-]+/g, ' ').toLowerCase();
        return norm.includes(baseCore);
      });
      if (match) return match.video_id;
      const total = res.data?.page_info?.total_number || 0;
      if (page * 20 >= total) break;
    } catch (_) { break; }
  }
  return null;
}

async function uploadVideos(urls, adv_id, jobId, rowIndex) {
  const ids = [];
  const coverPromises = {};
  const urlToId = {}; // cache URL→video_id within this upload batch
  for (let i = 0; i < urls.length; i++) {
    const url = urls[i];
    const videoNum = i + 1;
    // Same URL repeated in sheet — reuse without re-uploading
    if (urlToId[url]) {
      ids.push(urlToId[url]);
      continue;
    }
    try {
      // Use the TAIL of the filename so V1/V2/V3 suffixes are preserved and unique
      const fullName = url.split('/').pop().replace(/\.[^.]+$/, '');
      const tail = fullName.slice(-50); // take the end, not the start
      const rand = Math.random().toString(36).slice(2, 8);
      const video_name = `${tail}_${rand}`;
      const res = await ttPost('/file/video/ad/upload/', {
        upload_type: 'UPLOAD_BY_URL',
        video_url: url,
        video_name,
        flaw_detect: true,
        auto_fix_enabled: true,
      }, adv_id);
      // Check for video_id in success response
      const video_id = res.data?.video_id || res.data?.[0]?.video_id;
      if (video_id) {
        ids.push(video_id);
        urlToId[url] = video_id;
        coverPromises[video_id] = getVideoCoverImageId(video_id, adv_id);
      } else if (res.code === 40911) {
        // TikTok detected duplicate — check if it returned the existing ID directly
        const dup_id = res.data?.video_id || res.data?.[0]?.video_id;
        if (dup_id) {
          ids.push(dup_id);
          urlToId[url] = dup_id;
          coverPromises[dup_id] = getVideoCoverImageId(dup_id, adv_id);
        } else {
          // Fall back to library search using the URL tail (unique per V1/V2/V3)
          const fullName = url.split('/').pop().replace(/\.[^.]+$/, '');
          const searchTail = fullName.slice(-50);
          const existing_id = await findExistingVideo(searchTail, adv_id);
          if (existing_id) {
            console.log(`Reusing existing video for ${url}: ${existing_id}`);
            ids.push(existing_id);
            urlToId[url] = existing_id;
            coverPromises[existing_id] = getVideoCoverImageId(existing_id, adv_id);
          } else {
            const warn = `Video ${videoNum} skipped (duplicate, not found in library): ${url}`;
            console.warn(warn);
            if (jobId) jobEmit(jobId, { type: 'step', rowIndex, step: `⚠ ${warn}` });
          }
        }
      } else {
        const warn = `Video ${videoNum} failed (code=${res.code}): ${res.message}`;
        console.warn(warn);
        if (jobId) jobEmit(jobId, { type: 'step', rowIndex, step: `⚠ ${warn}` });
      }
    } catch (e) {
      const warn = `Video ${videoNum} error: ${e.message}`;
      console.warn(warn);
      if (jobId) jobEmit(jobId, { type: 'step', rowIndex, step: `⚠ ${warn}` });
    }
  }
  if (!ids.length) throw new Error('No videos uploaded successfully');
  return { ids, coverPromises };
}

async function createAdGroup(row, campaign_id, adv_id, pixel_id) {
  const location_id = GEO_MAP[row.geo];
  if (!location_id) throw new Error(`Unknown geo: ${row.geo}`);

  const rawDate = (row.start_date || '').toString().trim();
  let datePart = rawDate;
  const slashMatch = rawDate.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    datePart = `${slashMatch[3]}-${slashMatch[2].padStart(2,'0')}-${slashMatch[1].padStart(2,'0')}`;
  }
  const rawTime = (row.start_time || '').toString().trim();
  let hh = '00', mm = '00';
  const timeMatch = rawTime.match(/(\d{1,2}):(\d{2})/);
  if (timeMatch) { hh = timeMatch[1].padStart(2,'0'); mm = timeMatch[2]; }
  const schedule_start = `${datePart} ${hh}:${mm}:00`;

  // Map language name → code using the same list from /tool/language/
  const LANGUAGE_NAME_TO_CODE = {
    'arabic':'ar','assamese':'as','haryanvi':'bgc','bihari':'bh','bengali':'bn',
    'czech':'cs','german':'de','greek':'el','english':'en','spanish':'es',
    'finnish':'fi','french':'fr','gujarati':'gu','hebrew':'he','hindi':'hi',
    'hungarian':'hu','indonesian':'id','italian':'it','japanese':'ja','khmer':'km',
    'kannada':'kn','korean':'ko','malayalam':'ml','marathi':'mr','malay':'ms',
    'dutch':'nl','oriya':'or','punjabi':'pa','polish':'pl','portuguese':'pt',
    'rajasthani':'raj','romanian':'ro','russian':'ru','swedish':'sv','tamil':'ta',
    'telugu':'te','thai':'th','turkish':'tr','ukrainian':'uk','vietnamese':'vi',
    'simplified chinese':'zh','traditional chinese':'zh-hant',
  };
  const langCode = row.language ? LANGUAGE_NAME_TO_CODE[row.language.toLowerCase()] : null;

  const targeting_spec = {
    location_ids: [location_id],
    ...(row.targeting === 'AGE_35_PLUS' ? { age_groups: ['AGE_35_44', 'AGE_45_54', 'AGE_55_100'] } : {}),
    ...(langCode ? { languages: [langCode] } : {}),
  };

  return ttPost('/smart_plus/adgroup/create/', {
    request_id: (Date.now() + 1).toString(),
    campaign_id,
    adgroup_name: `${row.campaign_name}_adgroup`,
    promotion_type: 'WEBSITE',
    placement_type: 'PLACEMENT_TYPE_NORMAL',
    placements: ['PLACEMENT_TIKTOK'],
    targeting_optimization_mode: 'MANUAL',
    targeting_spec,
    optimization_goal: 'CONVERT',
    billing_event: 'OCPM',
    pixel_id,
    optimization_event: 'SHOPPING',
    schedule_type: 'SCHEDULE_FROM_NOW',
    schedule_start_time: schedule_start,
    bid_type: row.bid_strategy === 'COST_CAP' ? 'BID_TYPE_CUSTOM' : 'BID_TYPE_NO_BID',
    ...(row.bid_strategy === 'COST_CAP' && row.bid_amount ? { conversion_bid_price: row.bid_amount } : {}),
  }, adv_id);
}

async function getVideoCoverImageId(video_id, adv_id = ADV_ID) {
  for (let attempt = 1; attempt <= 20; attempt++) {
    try {
      const res = await ttGet('/file/video/suggestcover/', { video_id, poster_number: 1 }, adv_id);
      const cover = res.data?.list?.[0];
      console.log(`suggestcover attempt ${attempt} for ${video_id}: code=${res.code} cover=${JSON.stringify(cover)}`);
      if (cover?.url) {
        const uploadRes = await ttPost('/file/image/ad/upload/', {
          upload_type: 'UPLOAD_BY_URL',
          image_url: cover.url,
          image_name: `cover_${video_id}_${Date.now()}`,
        }, adv_id);
        const image_id = uploadRes.data?.image_id;
        if (image_id) { console.log(`Cover image_id for ${video_id}: ${image_id}`); return image_id; }
        if (uploadRes.code === 40911) {
          // TikTok detected duplicate image — check if it returned the existing ID directly
          const dup_image_id = uploadRes.data?.image_id;
          if (dup_image_id) { console.log(`Reusing duplicate cover image_id for ${video_id}: ${dup_image_id}`); return dup_image_id; }
          // Last resort: search by image URL hash to find this specific cover
          const searchRes = await ttGet('/file/image/ad/search/', {
            filtering: JSON.stringify({ image_name: `cover_${video_id}` }),
            page_size: 5,
          }, adv_id);
          const existing = searchRes.data?.list?.[0];
          if (existing?.image_id) { console.log(`Found cover by name for ${video_id}: ${existing.image_id}`); return existing.image_id; }
        }
        console.warn(`Cover upload failed: code=${uploadRes.code} msg=${uploadRes.message}`);
      }
    } catch (e) {
      console.warn(`suggestcover attempt ${attempt} error for ${video_id}:`, e.message);
    }
    await new Promise(r => setTimeout(r, 10000));
  }
  console.warn(`getVideoCoverImageId exhausted for ${video_id}`);
  return null;
}

async function createAds(row, adgroup_id, video_ids, identity_id, identity_type, identity_bc_id, coverPromises, adv_id) {
  const dedupedIds = [...new Set(video_ids)];
  const coverMap = {};
  for (const video_id of dedupedIds) {
    const image_id = await (coverPromises[video_id] || getVideoCoverImageId(video_id, adv_id));
    if (image_id) coverMap[video_id] = image_id;
  }

  const resolvedIdentityType = identity_type || 'BC_AUTH_TT';
  const creativeIdentity = {
    identity_type: resolvedIdentityType,
    identity_id,
    ...(resolvedIdentityType === 'BC_AUTH_TT' ? { identity_authorized_bc_id: identity_bc_id || BC_ID } : {}),
  };

  const creative_list = dedupedIds.map(video_id => {
    const cover = coverMap[video_id];
    if (!cover) throw new Error(`Could not fetch cover image for video ${video_id} — required for Smart Plus ads`);
    return {
      creative_info: {
        ad_format: 'SINGLE_VIDEO',
        video_info: { video_id },
        image_info: [{ web_uri: cover }],
        aigc_disclosure_type: 'SELF_DISCLOSURE',
        ...creativeIdentity,
      },
    };
  });

  const ad_text_list = row.headlines.slice(0, 5).map(h => ({ ad_text: h }));

  // call_to_action_list not supported with BC_AUTH_TT + TikTok placement; use portfolio IDs
  const CTA_PORTFOLIO_IDS = {
    LEARN_MORE: '7654255502322404372',
    SHOP_NOW: '7654256510972791828',
  };
  const call_to_action_id = CTA_PORTFOLIO_IDS.LEARN_MORE;

  for (let i = 0; i < creative_list.length; i += 50) {
    const batch = creative_list.slice(i, i + 50);
    const res = await ttPost('/smart_plus/ad/create/', {
      adgroup_id,
      ad_name: `${row.campaign_name}_ad_${Math.floor(i / 50) + 1}`,
      ad_configuration: { ...creativeIdentity, call_to_action_id },
      ad_text_list,
      landing_page_url_list: [{ landing_page_url: row.url }],
      creative_list: batch,
    }, adv_id);
    if (res.code !== 0) throw new Error(`Ad create failed: ${JSON.stringify(res)}`);
  }
}

// ── Start ──────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`Campaign launcher running at http://localhost:${PORT}`));
