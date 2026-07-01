require('dotenv').config();
const express = require('express');
const session = require('express-session');
const multer = require('multer');
const { parse } = require('csv-parse/sync');
const { v4: uuidv4 } = require('uuid');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(session({
  secret: process.env.SESSION_SECRET || 'secret',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 8 * 60 * 60 * 1000 }, // 8 hours
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

// Token stored in memory — survives until restart, then re-auth needed
let activeToken = process.env.TIKTOK_ACCESS_TOKEN || null;

function getToken() { return activeToken; }

async function ttGet(path, params = {}) {
  const url = new URL(`${TT_BASE}${path}`);
  url.searchParams.set('advertiser_id', ADV_ID);
  Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, typeof v === 'object' ? JSON.stringify(v) : v));
  const res = await fetch(url.toString(), { headers: { 'Access-Token': getToken() } });
  return res.json();
}

async function ttPost(path, body) {
  const res = await fetch(`${TT_BASE}${path}`, {
    method: 'POST',
    headers: { 'Access-Token': getToken(), 'Content-Type': 'application/json' },
    body: JSON.stringify({ advertiser_id: ADV_ID, ...body }),
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

// User pastes the auth_code from the advertising.tech redirect URL
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

// ── In-memory job store ────────────────────────────────────────────────────
const jobs = new Map(); // jobId -> { rows: [], status: 'running'|'done', clients: Set }

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

app.post('/api/logout', (req, res) => {
  req.session.destroy();
  res.json({ ok: true });
});

app.get('/api/me', (req, res) => {
  res.json({ user: req.session.user || null });
});

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

app.get('/sample', (req, res) => {
  const headers = [
    'campaign_name','geo','budget','bid_strategy','bid_amount','targeting',
    'start_date','start_time',
    'video_url_1','video_url_2','video_url_3','video_url_4','video_url_5',
    'video_url_6','video_url_7','video_url_8','video_url_9','video_url_10',
    'headline_1','headline_2','headline_3','headline_4','headline_5',
    'cta','url',
  ];
  const rows = [
    {
      campaign_name: 'BOZV_Jul30', geo: 'US', budget: 30,
      bid_strategy: 'LOWEST_COST', bid_amount: '', targeting: 'BROAD',
      start_date: '10/07/2026', start_time: '00:00',
      video_url_1: 'https://videosapi.net/videos/example1.mp4',
      video_url_2: 'https://videosapi.net/videos/example2.mp4',
      video_url_3: '', video_url_4: '', video_url_5: '',
      video_url_6: '', video_url_7: '', video_url_8: '',
      video_url_9: '', video_url_10: '',
      headline_1: 'Lose weight fast', headline_2: 'Try it free today',
      headline_3: 'Results in 7 days', headline_4: '', headline_5: '',
      cta: 'LEARN_MORE', url: 'https://yoursite.com/landing',
    },
    {
      campaign_name: 'MT_Jul30', geo: 'US', budget: 50,
      bid_strategy: 'COST_CAP', bid_amount: 15, targeting: 'AGE_35_PLUS',
      start_date: '10/07/2026', start_time: '08:00',
      video_url_1: 'https://videosapi.net/videos/example3.mp4',
      video_url_2: '', video_url_3: '', video_url_4: '', video_url_5: '',
      video_url_6: '', video_url_7: '', video_url_8: '',
      video_url_9: '', video_url_10: '',
      headline_1: 'Save more today', headline_2: 'Start saving now',
      headline_3: '', headline_4: '', headline_5: '',
      cta: 'SHOP_NOW', url: 'https://yoursite.com/offer',
    },
  ];

  const escape = v => {
    const s = String(v ?? '');
    return s.includes(',') || s.includes('"') || s.includes('\n') ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const csv = [headers, ...rows.map(r => headers.map(h => escape(r[h])))].map(r => r.join(',')).join('\r\n');
  res.setHeader('Content-Disposition', 'attachment; filename="campaign-template.csv"');
  res.setHeader('Content-Type', 'text/csv');
  res.send(csv);
});

app.post('/api/parse-csv', requireAuth, upload.single('csv'), (req, res) => {
  try {
    const text = req.file.buffer.toString('utf8');
    const records = parse(text, { columns: true, skip_empty_lines: true, trim: true });

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
      return {
        rowIndex: i,
        campaign_name: r.campaign_name,
        geo: (r.geo || '').toUpperCase(),
        budget: parseFloat(r.budget),
        bid_strategy: (r.bid_strategy || 'LOWEST_COST').toUpperCase(),
        bid_amount: parseFloat(r.bid_amount) || null,
        targeting: (r.targeting || 'BROAD').toUpperCase(),
        start_date: r.start_date,
        start_time: r.start_time || '00:00',
        videos,
        headlines,
        cta: r.cta || 'LEARN_MORE',
        url: r.url,
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
  const { rows, identity_id, identity_type, identity_bc_id } = req.body;
  if (!rows?.length) return res.status(400).json({ error: 'No rows' });

  const jobId = uuidv4();
  jobs.set(jobId, { events: [], clients: new Set(), status: 'running' });
  res.json({ jobId });

  // Run deployment in background
  deployRows(jobId, rows, identity_id, identity_type, identity_bc_id).catch(console.error);
});

app.get('/api/deploy/:jobId/events', requireAuth, (req, res) => {
  const job = jobs.get(req.params.jobId);
  if (!job) return res.status(404).end();

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  // Send buffered events first
  job.events.forEach(e => res.write(`data: ${JSON.stringify(e)}\n\n`));

  if (job.status === 'done') return res.end();

  job.clients.add(res);
  req.on('close', () => job.clients.delete(res));
});

// ── Deployment logic ───────────────────────────────────────────────────────
async function deployRows(jobId, rows, identity_id, identity_type, identity_bc_id) {
  for (const row of rows) {
    jobEmit(jobId, { type: 'row_start', rowIndex: row.rowIndex, campaign_name: row.campaign_name });

    try {
      // 1. Create campaign
      jobEmit(jobId, { type: 'step', rowIndex: row.rowIndex, step: 'Creating campaign…' });
      const campaign = await createCampaign(row);
      const campaign_id = campaign.data?.campaign_id;
      if (!campaign_id) {
        const msg = campaign.message || JSON.stringify(campaign);
        throw new Error(campaign.code === 40911 || (msg && msg.includes('already')) ? `Campaign name "${row.campaign_name}" already exists in TikTok — rename it in the CSV and retry` : `Campaign failed: ${msg}`);
      }

      // 2. Upload videos
      jobEmit(jobId, { type: 'step', rowIndex: row.rowIndex, step: `Uploading ${row.videos.length} video(s)…` });
      const video_ids = await uploadVideos(row.videos);

      // 3. Create ad group
      jobEmit(jobId, { type: 'step', rowIndex: row.rowIndex, step: 'Creating ad group…' });
      const adgroup = await createAdGroup(row, campaign_id);
      const adgroup_id = adgroup.data?.adgroup_id;
      if (!adgroup_id) throw new Error(`Ad group failed: ${JSON.stringify(adgroup)}`);

      // 4. Create ads (video × headline combinations)
      jobEmit(jobId, { type: 'step', rowIndex: row.rowIndex, step: `Creating ${video_ids.length * row.headlines.length} ads…` });
      await createAds(row, adgroup_id, video_ids, identity_id, identity_type, identity_bc_id);

      jobEmit(jobId, { type: 'row_done', rowIndex: row.rowIndex, campaign_id });

    } catch (err) {
      jobEmit(jobId, { type: 'row_error', rowIndex: row.rowIndex, error: err.message });
    }
  }

  jobEmit(jobId, { type: 'done' });
  const job = jobs.get(jobId);
  if (job) {
    job.status = 'done';
    job.clients.forEach(r => r.end());
  }
}

async function createCampaign(row) {
  const body = {
    campaign_name: row.campaign_name,
    objective_type: 'WEB_CONVERSIONS',
    virtual_objective_type: 'SALES',
    sales_destination: 'WEBSITE',
    budget_optimize_on: true,
    budget_mode: 'BUDGET_MODE_DYNAMIC_DAILY_BUDGET',
    budget: row.budget,
    operation_status: 'ENABLE',
    campaign_automation_type: 'UPGRADED_SMART_PLUS',
  };
  return ttPost('/campaign/create/', body);
}

async function findExistingVideo(baseName) {
  // Strip leading numeric prefix and normalize for matching
  const baseCore = baseName.replace(/^\d+_/, '').replace(/[_\s-]+/g, ' ').toLowerCase().slice(0, 25);
  for (let page = 1; page <= 15; page++) {
    try {
      const res = await ttGet('/file/video/ad/search/', { page, page_size: 20 });
      const list = res.data?.list || [];
      const match = list.find(v => {
        if (!v.file_name) return false;
        return v.file_name.replace(/[_\s-]+/g, ' ').toLowerCase().includes(baseCore);
      });
      if (match) return match.video_id;
      const total = res.data?.page_info?.total_number || 0;
      if (page * 20 >= total) break;
    } catch (e) {
      break;
    }
  }
  return null;
}

async function uploadVideos(urls) {
  const ids = [];
  const errors = [];
  for (const url of urls) {
    try {
      const baseName = url.split('/').pop().replace(/\.[^.]+$/, '').slice(0, 40);
      const rand = Math.random().toString(36).slice(2, 8);
      const video_name = `${baseName}_${Date.now()}_${rand}`;
      const res = await ttPost('/file/video/ad/upload/', {
        upload_type: 'UPLOAD_BY_URL',
        video_url: url,
        video_name,
        flaw_detect: true,
        auto_fix_enabled: true,
      });
      const video_id = res.data?.video_id || res.data?.[0]?.video_id;
      if (video_id) {
        ids.push(video_id);
      } else if (res.code === 40911) {
        // Video already exists in library — find and reuse it
        const existing_id = await findExistingVideo(baseName);
        if (existing_id) {
          console.log(`Reusing existing video for ${url}: ${existing_id}`);
          ids.push(existing_id);
        } else {
          const msg = `Upload failed (duplicate) and could not find existing video for ${url}`;
          console.warn(msg);
          errors.push(msg);
        }
      } else {
        const msg = `Upload failed for ${url}: code=${res.code} msg=${res.message} data=${JSON.stringify(res.data)}`;
        console.warn(msg);
        errors.push(msg);
      }
    } catch (e) {
      const msg = `Upload error for ${url}: ${e.message}`;
      console.warn(msg);
      errors.push(msg);
    }
  }
  if (!ids.length) throw new Error(errors.join(' | ') || 'No videos uploaded successfully');
  return ids;
}

async function createAdGroup(row, campaign_id) {
  const location_id = GEO_MAP[row.geo];
  if (!location_id) throw new Error(`Unknown geo: ${row.geo}`);

  // Parse date: handles YYYY-MM-DD or DD/MM/YYYY or MM/DD/YYYY -> YYYY-MM-DD
  const rawDate = (row.start_date || '').toString().trim();
  let datePart = rawDate;
  const slashMatch = rawDate.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    // Treat as DD/MM/YYYY (Excel default in many locales)
    datePart = `${slashMatch[3]}-${slashMatch[2].padStart(2,'0')}-${slashMatch[1].padStart(2,'0')}`;
  }
  // Parse time: handles "8:00", "08:00", blank -> "00:00"
  const rawTime = (row.start_time || '').toString().trim();
  let hh = '00', mm = '00';
  const timeMatch = rawTime.match(/(\d{1,2}):(\d{2})/);
  if (timeMatch) { hh = timeMatch[1].padStart(2,'0'); mm = timeMatch[2]; }
  const schedule_start = `${datePart} ${hh}:${mm}:00`;

  const body = {
    campaign_id,
    adgroup_name: `${row.campaign_name}_adgroup`,
    placement_type: 'PLACEMENT_TYPE_NORMAL',
    placements: ['PLACEMENT_TIKTOK'],
    budget_mode: 'BUDGET_MODE_INFINITE',
    schedule_type: 'SCHEDULE_FROM_NOW',
    schedule_start_time: schedule_start,
    optimization_goal: 'CONVERT',
    billing_event: 'OCPM',
    pacing: 'PACING_MODE_SMOOTH',
    promotion_type: 'WEBSITE',
    pixel_id: PIXEL_ID,
    optimization_event: 'SHOPPING',
    location_ids: [location_id],
    bid_type: row.bid_strategy === 'COST_CAP' ? 'BID_TYPE_CUSTOM' : 'BID_TYPE_NO_BID',
    ...(row.bid_strategy === 'COST_CAP' && row.bid_amount ? { conversion_bid_price: row.bid_amount } : {}),
    ...(row.targeting === 'AGE_35_PLUS' ? { age_groups: ['AGE_35_44', 'AGE_45_54', 'AGE_55_100'] } : {}),
  };

  return ttPost('/adgroup/create/', body);
}

async function getVideoCoverId(video_id) {
  try {
    const res = await ttGet('/file/video/suggestcover/get/', { video_id, poster_number: 1 });
    const cover = res.data?.list?.[0];
    if (cover?.image_id) return cover.image_id;
  } catch (e) {
    console.warn(`Could not fetch cover for ${video_id}:`, e.message);
  }
  return null;
}

async function createAds(row, adgroup_id, video_ids, identity_id, identity_type, identity_bc_id) {
  // Pre-fetch a cover image ID for each video (required for non-Spark SINGLE_VIDEO ads)
  const coverMap = {};
  for (const video_id of video_ids) {
    const cover_id = await getVideoCoverId(video_id);
    if (cover_id) coverMap[video_id] = cover_id;
  }

  const creatives = [];
  for (const video_id of video_ids) {
    for (const headline of row.headlines) {
      creatives.push({
        ad_name: `${row.campaign_name}_${video_id.slice(-6)}_${creatives.length + 1}`,
        ad_format: 'SINGLE_VIDEO',
        identity_type,
        identity_id,
        ...(identity_type === 'BC_AUTH_TT' ? { identity_authorized_bc_id: identity_bc_id } : {}),
        video_id,
        ...(coverMap[video_id] ? { image_ids: [coverMap[video_id]] } : {}),
        ad_text: headline,
        display_name: row.campaign_name.slice(0, 40),
        call_to_action: row.cta,
        landing_page_url: row.url,
      });
    }
  }

  // TikTok allows max 20 creatives per ad/create call
  for (let i = 0; i < creatives.length; i += 20) {
    const batch = creatives.slice(i, i + 20);
    const res = await ttPost('/ad/create/', { adgroup_id, creatives: batch });
    if (res.code !== 0) throw new Error(`Ad create failed: ${JSON.stringify(res)}`);
  }
}

// ── Start ──────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`Campaign launcher running at http://localhost:${PORT}`));
