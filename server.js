const express = require('express');
const fetch = require('node-fetch');
const path = require('path');
const XLSX = require('xlsx');
const mammoth = require('mammoth');
const fs = require('fs');
const { google } = require('googleapis');
const { Pool }   = require('pg');
const app = require('express')();

// ─── POSTGRESQL (FIELD LOG) ───────────────────────────────────────────────────
const pgPool = process.env.DATABASE_URL
  ? new Pool({
      connectionString:      process.env.DATABASE_URL,
      ssl:                   { rejectUnauthorized: false },
      max:                   10,
      idleTimeoutMillis:     30000,
      connectionTimeoutMillis: 2000,
    })
  : null;

if (pgPool) {
  pgPool.on('error', (err) => {
    console.error('Unexpected error on idle pg client', err.message);
  });
}

async function initFieldLogTable() {
  if (!pgPool) return;
  await pgPool.query(`
    CREATE TABLE IF NOT EXISTS field_log_entries (
      id             SERIAL PRIMARY KEY,
      rep_name       TEXT NOT NULL,
      logged_at      TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      pm_opp         BOOLEAN DEFAULT FALSE,
      equip_opp      BOOLEAN DEFAULT FALSE,
      service_lead   BOOLEAN DEFAULT FALSE,
      piping_opp     BOOLEAN DEFAULT FALSE,
      sticker        BOOLEAN DEFAULT FALSE,
      vr_pres        BOOLEAN DEFAULT FALSE,
      appt_set       BOOLEAN DEFAULT FALSE,
      nothing        BOOLEAN DEFAULT FALSE,
      location       TEXT,
      notable_moment TEXT
    )
  `);
  await pgPool.query(`ALTER TABLE field_log_entries ADD COLUMN IF NOT EXISTS company_name  TEXT`);
  await pgPool.query(`ALTER TABLE field_log_entries ADD COLUMN IF NOT EXISTS contact_name  TEXT`);
  await pgPool.query(`ALTER TABLE field_log_entries ADD COLUMN IF NOT EXISTS sticker_count INTEGER DEFAULT 0`);
  await pgPool.query(`ALTER TABLE field_log_entries ADD COLUMN IF NOT EXISTS mobile        TEXT`);
  await pgPool.query(`ALTER TABLE field_log_entries ADD COLUMN IF NOT EXISTS office_phone  TEXT`);
  await pgPool.query(`ALTER TABLE field_log_entries ADD COLUMN IF NOT EXISTS email         TEXT`);
  await pgPool.query(`ALTER TABLE field_log_entries ADD COLUMN IF NOT EXISTS website       TEXT`);
  await pgPool.query(`ALTER TABLE field_log_entries ADD COLUMN IF NOT EXISTS card_address    TEXT`);
  await pgPool.query(`ALTER TABLE field_log_entries ADD COLUMN IF NOT EXISTS contact_title  TEXT`);
  await pgPool.query(`CREATE INDEX IF NOT EXISTS idx_field_log_rep_logged ON field_log_entries (rep_name, logged_at DESC)`);
  await pgPool.query(`
    CREATE TABLE IF NOT EXISTS rep_goals (
      rep_name       TEXT PRIMARY KEY,
      wednesday_goal INTEGER NOT NULL DEFAULT 25,
      daily_goal     INTEGER,
      weekly_goal    INTEGER NOT NULL DEFAULT 30,
      updated_at     TIMESTAMPTZ DEFAULT NOW()
    )
  `);
  console.log('field_log_entries table ready');
}

async function initBlitzTables() {
  if (!pgPool) return;
  await pgPool.query(`
    CREATE TABLE IF NOT EXISTS blitz_sessions (
      id SERIAL PRIMARY KEY,
      name TEXT NOT NULL,
      created_at TIMESTAMPTZ DEFAULT NOW(),
      active BOOLEAN DEFAULT TRUE
    )
  `);
  await pgPool.query(`
    CREATE TABLE IF NOT EXISTS blitz_stops (
      id SERIAL PRIMARY KEY,
      session_id INTEGER REFERENCES blitz_sessions(id),
      group_name TEXT NOT NULL,
      day INTEGER NOT NULL,
      stop_number TEXT NOT NULL,
      role TEXT NOT NULL,
      company TEXT NOT NULL,
      address TEXT,
      miles_from_shop NUMERIC,
      industry TEXT,
      contact_name TEXT,
      title TEXT,
      phone TEXT,
      air_likelihood TEXT,
      site_confidence TEXT,
      territory TEXT,
      priority_score NUMERIC,
      website TEXT,
      contact_email TEXT,
      logged BOOLEAN DEFAULT FALSE,
      logged_at TIMESTAMPTZ,
      outcome TEXT,
      notes TEXT
    )
  `);
  await pgPool.query(`CREATE INDEX IF NOT EXISTS idx_blitz_stops_session ON blitz_stops (session_id)`);
  await pgPool.query(`ALTER TABLE blitz_sessions ADD COLUMN IF NOT EXISTS group_assignments JSONB DEFAULT '{}'`);
  await pgPool.query(`ALTER TABLE blitz_stops ADD COLUMN IF NOT EXISTS email TEXT`);
  await pgPool.query(`ALTER TABLE blitz_stops ADD COLUMN IF NOT EXISTS mobile TEXT`);
  console.log('blitz tables ready');
}

// ─── BRANCH TIMEZONE MAP ─────────────────────────────────────────────────────
const BRANCH_TZ = {
  Atlanta:    'America/New_York',
  Charlotte:  'America/New_York',
  Tampa:      'America/New_York',
  Greenville: 'America/New_York',
  Cleveland:  'America/New_York',
  Nashville:  'America/Chicago',
  Detroit:    'America/Chicago',
  Dallas:     'America/Chicago',
  Chicago:    'America/Chicago',
};
// branch values from client are e.g. "Nashville Compressor" — match by substring
function getTZ(branch) {
  if (!branch) return 'America/New_York';
  const key = Object.keys(BRANCH_TZ).find(k => branch.toLowerCase().includes(k.toLowerCase()));
  return key ? BRANCH_TZ[key] : 'America/New_York';
}

// ─── ESTIMATES PERSISTENCE ────────────────────────────────────────────────────
const ESTIMATES_FILE = path.join(__dirname, 'data', 'estimates.json');

function ensureEstimatesFile() {
  const dir = path.join(__dirname, 'data');
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  if (!fs.existsSync(ESTIMATES_FILE)) fs.writeFileSync(ESTIMATES_FILE, '[]', 'utf8');
}

function readEstimates() {
  ensureEstimatesFile();
  try { return JSON.parse(fs.readFileSync(ESTIMATES_FILE, 'utf8')); } catch { return []; }
}

function writeEstimates(list) {
  ensureEstimatesFile();
  fs.writeFileSync(ESTIMATES_FILE, JSON.stringify(list, null, 2), 'utf8');
}

const STORIES_FILE = path.join(__dirname, 'data', 'stories.json');

function ensureStoriesFile() {
  const dir = path.join(__dirname, 'data');
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  if (!fs.existsSync(STORIES_FILE)) fs.writeFileSync(STORIES_FILE, '[]', 'utf8');
}

function readStories() {
  ensureStoriesFile();
  try { return JSON.parse(fs.readFileSync(STORIES_FILE, 'utf8')); } catch { return []; }
}

function writeStories(list) {
  ensureStoriesFile();
  fs.writeFileSync(STORIES_FILE, JSON.stringify(list, null, 2), 'utf8');
}

app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static('.'));

const PROXY_URL = 'https://airiq-st-proxy.jon-sanders.workers.dev';

// ─── GOOGLE SHEETS CONFIG ─────────────────────────────────────────────────────
const SHEET_ID   = '1Zrzr_PgEMmMmclJ2W1Dere-JEKYmaB4xa-N4MMUIUuI';
const SHEET_NAME = 'Sheet1';
const SHEET_HEADERS = [
  'Date','Rep','Branch','Customer','Contact','Drive Time RT','Equipment Delivery',
  'Electrical By','Piping By','New Equipment','Existing Equipment','Equipment Location',
  'Loading Dock','Forklift','Unloading Issues','Path Clear','Doors/Halls',
  'Levelers Required','Disconnect Within 10ft','Breaker Space','Wire Size',
  'Electrical Notes','Pipe Material','Pipe Size','Linear Feet','Pipe Height Over 8ft',
  'Elbows','Tees','Couplings','Ball Valves','NPT Connectors','Bushings/Reducers',
  'Unions','Pipe Clips','Unistrut','Quick Branch','Bypasses','Drops','Drop Length/Size',
  'Termination Blocks','Quick Connects','Obstructions','Beams/Workarounds','Piping Notes',
  'Diesel Compressor','Scissor Lift','Forklift Rental','Lugging Tool','Other Rental',
  'Drawing Checklist Complete','Manager Approval','Notes','Submitted At','Site Drawing',
];

// ─── RETRY HELPER ────────────────────────────────────────────────────────────
async function retryFetch(fn, retries = 3, delays = [1000, 5000, 30000]) {
  let lastErr;
  for (let attempt = 0; attempt <= retries; attempt++) {
    try { return await fn(); } catch (err) {
      lastErr = err;
      if (attempt < retries) {
        console.log(`Retrying attempt ${attempt + 2} of ${retries + 1}...`);
        await new Promise(r => setTimeout(r, delays[attempt] ?? delays[delays.length - 1]));
      }
    }
  }
  throw lastErr;
}

// In-memory inventory cache
let inventoryCache = [];
let lastCacheUpdate = null;
let isFetching = false;
let inventoryFromFallback = false;
let fallbackCacheDate = null;

const INVENTORY_CACHE_FILE = path.join('/app/data', 'inventory_cache.json');

function writeInventoryFallback(items) {
  try {
    const dir = '/app/data';
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(INVENTORY_CACHE_FILE, JSON.stringify({ items, writtenAt: new Date().toISOString() }), 'utf8');
    console.log('Inventory fallback cache written:', items.length, 'rows');
  } catch (e) { console.error('Failed to write inventory fallback cache:', e.message); }
}

function loadInventoryFallback() {
  try {
    if (!fs.existsSync(INVENTORY_CACHE_FILE)) return false;
    const { items, writtenAt } = JSON.parse(fs.readFileSync(INVENTORY_CACHE_FILE, 'utf8'));
    if (items && items.length > 0) {
      inventoryCache = items;
      lastCacheUpdate = writtenAt;
      inventoryFromFallback = true;
      fallbackCacheDate = writtenAt;
      console.log('Inventory loaded from fallback cache, rows:', items.length);
      return true;
    }
  } catch (e) { console.error('Failed to load inventory fallback cache:', e.message); }
  return false;
}

// ─── GMAIL INVENTORY LOADER ───────────────────────────────────────────────────
// Reads the daily "Hodge Compressor Inventory" email from Gmail,
// downloads the .xlsx attachment, and populates the inventory cache.
// No ST API calls needed — zero rate limit risk.

async function getGmailToken() {
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Accept-Encoding': 'identity' },
    compress: false,
    body: new URLSearchParams({
      client_id:     process.env.GMAIL_CLIENT_ID,
      client_secret: process.env.GMAIL_CLIENT_SECRET,
      refresh_token: process.env.GMAIL_REFRESH_TOKEN,
      grant_type:    'refresh_token',
    }),
  });
  const tokenText = await res.text();
  const data = JSON.parse(tokenText);
  if (!data.access_token) {
    throw new Error('Gmail token error: ' + JSON.stringify(data));
  }
  return data.access_token;
}

// ─── GOOGLE SHEETS ────────────────────────────────────────────────────────────

function buildSheetRow(body) {
  const d = body.data || {};
  const yn = v => v === true ? 'Yes' : v === false ? 'No' : '';
  const newEq = [
    d.eqRotaryScrew  && 'Rotary Screw',
    d.eqRecip        && 'Reciprocating',
    d.eqDryer        && 'Dryer',
    d.eqReceiver     && 'Receiver Tank',
    d.eqFilter       && 'Filter',
    d.eqOWSep        && 'O/W Separator',
    d.eqAutoDrains   && 'Auto Drains',
    d.eqSafetyValve  && 'Safety Valve',
    d.eqOther        && (d.eqOtherText || 'Other'),
  ].filter(Boolean).join(', ');
  const drawDone = [
    d.drawDrops, d.drawElectrical, d.drawBypasses, d.drawEquipment,
    d.drawBeams, d.drawDock, d.drawCustomerEquip, d.drawPiping, d.drawDiagrams,
  ].every(Boolean) ? 'Yes' : 'No';
  const withNote = (flag, note) => flag != null
    ? (yn(flag) + (note ? ': ' + note : ''))
    : '';
  return [
    d.date || '', d.rep || '', d.branch || '', d.customer || '', d.contact || '',
    d.driveTime || '', d.delivery || '', d.electrical || '', d.piping || '',
    newEq, d.existingEquip || '', d.equipLocation || '',
    yn(d.hasDock), yn(d.hasForklift),
    withNote(d.unloadingIssues, d.unloadingNotes),
    yn(d.clearPath),
    withNote(d.doorsHalls, d.doorsHallsNotes),
    yn(d.needLevelers), yn(d.fusDisconnect), yn(d.circuitBreaker),
    d.wireSize || '', d.electricalNotes || '',
    d.pipeMaterial || '', d.pipeSize || '', d.linearFeet || '',
    yn(d.pipeAbove8ft),
    d.elbows90 || '', d.tees || '', d.couplings || '', d.ballValves || '',
    d.nptConnectors || '', d.bushings || '', d.unions || '',
    d.pipeClips || '', d.unistrut || '', d.quickBranch || '',
    d.bypasses || '', d.numDrops || '', d.dropLength || '',
    d.terminationBlocks || '', d.quickConnects || '',
    withNote(d.equipObstructions, d.equipObstructionsNotes),
    withNote(d.beams, d.beamsNotes),
    d.pipingNotes || '',
    d.rentalDiesel   ? 'Yes' : 'No',
    d.rentalScissor  ? 'Yes' : 'No',
    d.rentalForklift ? 'Yes' : 'No',
    d.rentalLugging  ? 'Yes' : 'No',
    d.rentalOther || '',
    drawDone, yn(d.managerApproval), d.approvalNotes || '',
    body.savedAt || new Date().toISOString(),
    d.drawingDataURL ? 'Yes (see record)' : '',
  ];
}

async function appendToSheet(rowData) {
  try {
    const token = await getGmailToken();
    const range = encodeURIComponent(`${SHEET_NAME}!A1`);
    const res = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${range}:append?valueInputOption=USER_ENTERED`,
      {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json', 'Accept-Encoding': 'identity' },
        body: JSON.stringify({ values: [rowData] }),
      }
    );
    const data = await res.json();
    if (!res.ok) { console.error('Sheets append error:', JSON.stringify(data)); return false; }
    return true;
  } catch (err) {
    console.error('Sheets append exception:', err.message);
    return false;
  }
}

async function fetchInventoryFromGmail() {
  if (isFetching) {
    console.log('Inventory fetch already in progress, skipping...');
    return;
  }
  isFetching = true;
  console.log('Loading inventory from Gmail...');

  try {
    console.log('Gmail: fetching OAuth token...');
    const oauth2Client = new google.auth.OAuth2(
      process.env.GMAIL_CLIENT_ID,
      process.env.GMAIL_CLIENT_SECRET
    );
    oauth2Client.setCredentials({ refresh_token: process.env.GMAIL_REFRESH_TOKEN });
    const gmail = google.gmail({ version: 'v1', auth: oauth2Client });
    console.log('Gmail: OAuth token obtained');

    // Step 1: Search for the most recent Hodge Compressor Inventory email
    console.log('Gmail: step 1 — searching for inventory email...');
    const searchRes = await gmail.users.messages.list({
      userId: 'me',
      maxResults: 1,
      q: 'subject:"Hodge Compressor Inventory" from:noreply@onservicetitan.com',
    });
    const messages = searchRes.data.messages;

    if (!messages || messages.length === 0) {
      console.error('No Hodge Compressor Inventory email found in Gmail');
      return;
    }

    const messageId = messages[0].id;
    console.log('Gmail: step 1 complete, messageId:', messageId);

    // Step 2: Get the message to find the attachment ID
    console.log('Gmail: step 2 — fetching message details...');
    const msgRes = await gmail.users.messages.get({
      userId: 'me',
      id: messageId,
    });
    const msgData = msgRes.data;

    // Find the xlsx attachment part
    let attachmentId = null;
    let filename = null;

    function findAttachment(parts) {
      if (!parts) return;
      for (const part of parts) {
        if (
          part.mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
          (part.filename && part.filename.endsWith('.xlsx'))
        ) {
          attachmentId = part.body.attachmentId;
          filename = part.filename;
          return;
        }
        if (part.parts) findAttachment(part.parts);
      }
    }

    findAttachment(msgData.payload.parts);

    if (!attachmentId) {
      console.error('No .xlsx attachment found in inventory email');
      return;
    }

    console.log('Gmail: step 2 complete, attachmentId:', attachmentId);

    // Step 3: Download the attachment
    console.log('Gmail: step 3 — downloading attachment...');
    const attachRes = await gmail.users.messages.attachments.get({
      userId: 'me',
      messageId,
      id: attachmentId,
    });
    const attachData = attachRes.data;
    console.log('Gmail: step 3 complete, attachment size:', attachData?.size);

    // Gmail returns base64url encoded data — convert to standard base64
    const base64 = attachData.data.replace(/-/g, '+').replace(/_/g, '/');
    const buffer = Buffer.from(base64, 'base64');

    // Step 4: Parse the xlsx
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Row 0 is the header — skip it
    const dataRows = rows.slice(1);

    // Column order from the ST report:
    // 0: Item Name, 1: Item Code, 2: Item Description, 3: Item Type,
    // 4: Inventory Location, 5: Bin Location, 6: Quantity Available,
    // 7: Quantity on Hold, 8: Quantity on Hand, 9: Quantity on Order,
    // 10: Min Quantity, 11: Max Quantity, 12: Quantity Reserved,
    // 13: Quantity Staged, 14: Quantity on Site, 15: Quantity Available to Pick,
    // 16: Quantity to Order

    const EQUIPMENT_PREFIXES = ['HD','HB','HV','HT','HAD','HTM'];

    const parsed = dataRows
      .filter(row => {
        const code = (row[1] || '').toString().toUpperCase();
        return EQUIPMENT_PREFIXES.some(p => code.startsWith(p));
      })
      .map(row => ({
        name:               row[0] || '',
        code:               row[1] || '',
        description:        row[2] || '',
        type:               row[3] || '',
        location:           row[4] || '',
        binLocation:        row[5] || '',
        quantityAvailable:  Number(row[6])  || 0,
        quantityOnHold:     Number(row[7])  || 0,
        quantityOnHand:     Number(row[8])  || 0,
        quantityOnOrder:    Number(row[9])  || 0,
        quantityReserved:   Number(row[12]) || 0,
        quantityStaged:     Number(row[13]) || 0,
        quantityOnSite:     Number(row[14]) || 0,
      }));

    inventoryCache = parsed;
    lastCacheUpdate = new Date().toISOString();
    inventoryFromFallback = false;
    fallbackCacheDate = null;
    if (parsed.length > 0) writeInventoryFallback(parsed);
    console.log(`Inventory cache loaded from Gmail: ${parsed.length} equipment items at ${lastCacheUpdate}`);

    if (parsed.length > 0) {
      console.log('Sample item:', JSON.stringify(parsed[0]));
    }

  } catch (err) {
    console.error('fetchInventoryFromGmail error:', err);
  } finally {
    isFetching = false;
  }
}

// ─── INVENTORY SEARCH & AGGREGATION ──────────────────────────────────────────

function searchInventoryCache(searchTerm) {
  const term = searchTerm.toUpperCase();
  return inventoryCache.filter(item =>
    (item.code && item.code.toUpperCase().startsWith(term)) ||
    (item.name && item.name.toUpperCase().startsWith(term))
  );
}

function aggregateInventory(rows) {
  if (rows.length === 0) return rows;

  const totals = {
    quantityOnHand: 0,
    quantityAvailable: 0,
    quantityOnOrder: 0,
    quantityReserved: 0,
    locations: []
  };

  rows.forEach(row => {
    totals.quantityOnHand    += (row.quantityOnHand    || 0);
    totals.quantityAvailable += (row.quantityAvailable || 0);
    totals.quantityOnOrder   += (row.quantityOnOrder   || 0);
    totals.quantityReserved  += (row.quantityReserved  || 0);
    if ((row.quantityOnHand || 0) > 0 || (row.quantityAvailable || 0) > 0) {
      totals.locations.push({
        location:  row.location,
        onHand:    row.quantityOnHand,
        available: row.quantityAvailable,
        onOrder:   row.quantityOnOrder,
      });
    }
  });

  return totals;
}

// ─── API ROUTES ───────────────────────────────────────────────────────────────

// Parse Word doc (.docx) and extract raw text
app.post('/api/parse-docx', async (req, res) => {
  try {
    const { base64 } = req.body;
    if (!base64) return res.status(400).json({ error: 'No base64 data provided' });
    const buffer = Buffer.from(base64, 'base64');
    const result = await mammoth.extractRawText({ buffer });
    res.json({ text: result.value });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Claude proxy
app.post('/api/claude', async (req, res) => {
  const claudeHeaders = {
    'Content-Type': 'application/json',
    'Accept': 'application/json',
    'Accept-Encoding': 'identity',
    'x-api-key': process.env.ANTHROPIC_API_KEY,
    'anthropic-version': '2023-06-01',
    'anthropic-beta': 'web-search-2025-03-05'
  };
  const claudeBody = JSON.stringify(req.body);
  async function attemptClaude() {
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: claudeHeaders,
      body: claudeBody,
    });
    return response.json();
  }
  try {
    const data = await retryFetch(() => attemptClaude());
    res.json(data);
  } catch (e) {
    console.error('claude fetch failed after retry:', e);
    res.status(500).json({ error: e.message });
  }
});

// ST API token cache
let stAccessToken = null;
let stTokenExpiry = 0;

async function stGetToken() {
  if (stAccessToken && Date.now() < stTokenExpiry - 60000) return stAccessToken;
  const { ST_CLIENT_ID, ST_CLIENT_SECRET } = process.env;
  if (!ST_CLIENT_ID || !ST_CLIENT_SECRET) throw new Error('ST credentials not configured');
  const credentials = Buffer.from(`${ST_CLIENT_ID}:${ST_CLIENT_SECRET}`).toString('base64');
  const resp = await fetch('https://auth.servicetitan.io/connect/token', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
      'Authorization': `Basic ${credentials}`,
      'Accept-Encoding': 'identity',
    },
    body: new URLSearchParams({ grant_type: 'client_credentials' }),
  });
  if (!resp.ok) {
    const body = await resp.text();
    throw new Error(`ST auth failed: ${resp.status} — ${body.slice(0, 200)}`);
  }
  const d = await resp.json();
  stAccessToken = d.access_token;
  stTokenExpiry = Date.now() + (d.expires_in || 3600) * 1000;
  return stAccessToken;
}

// ST customer lookup
// NOTE: makes direct calls to api.servicetitan.io — the CF Worker (PROXY_URL) is pricebook-only
// and would need a customer-search handler added before this could route through it.
// Requires Railway env vars: ST_TENANT, ST_CLIENT_ID, ST_CLIENT_SECRET
app.get('/api/st/customers', async (req, res) => {
  const name = (req.query.name || '').trim();
  if (name.length < 2) return res.json({ customers: [] });
  const { ST_TENANT, ST_CLIENT_ID, ST_CLIENT_SECRET, ST_APP_KEY } = process.env;
  console.log(`[ST customers] query="${name}" tenant=${ST_TENANT ? ST_TENANT.slice(0,6)+'…' : 'MISSING'} clientId=${ST_CLIENT_ID ? 'SET' : 'MISSING'} secret=${ST_CLIENT_SECRET ? 'SET' : 'MISSING'} appKey=${ST_APP_KEY ? 'SET' : 'MISSING'}`);
  if (!ST_TENANT || !ST_CLIENT_ID || !ST_CLIENT_SECRET) return res.json({ notConfigured: true, customers: [] });
  try {
    const token = await stGetToken();
    console.log(`[ST customers] token acquired (${token.length} chars)`);
    const appKeyParam = ST_APP_KEY ? `&appKey=${encodeURIComponent(ST_APP_KEY)}` : '';
    const hdrs = { 'Authorization': `Bearer ${token}`, 'ST-App-Key': ST_APP_KEY || ST_CLIENT_ID };

    const custUrl = `https://api.servicetitan.io/crm/v2/tenant/${ST_TENANT}/customers?name=${encodeURIComponent(name)}&active=true&pageSize=5${appKeyParam}`;
    console.log(`[ST customers] GET ${custUrl}`);
    const custResp = await fetch(custUrl, { headers: { ...hdrs, 'Accept-Encoding': 'identity' } });
    if (!custResp.ok) {
      const body = await custResp.text();
      console.error(`[ST customers] customer search failed status=${custResp.status} body=${body}`);
      throw new Error(`ST customers: ${custResp.status} — ${body.slice(0, 200)}`);
    }
    const custData = await custResp.json();
    console.log(`[ST customers] found ${(custData.data || []).length} results`);
    const customers = (custData.data || []).slice(0, 5);

    let lastServiceDate = null;
    if (customers.length > 0) {
      try {
        const jobsUrl = `https://api.servicetitan.io/jpm/v2/tenant/${ST_TENANT}/jobs?customerId=${customers[0].id}&pageSize=1&sort=-completedOn&jobStatus=Completed${appKeyParam}`;
        console.log(`[ST customers] GET ${jobsUrl}`);
        const jobsResp = await fetch(jobsUrl, { headers: { ...hdrs, 'Accept-Encoding': 'identity' } });
        if (jobsResp.ok) {
          const jd = await jobsResp.json();
          lastServiceDate = (jd.data || [])[0]?.completedOn || null;
          console.log(`[ST customers] lastServiceDate=${lastServiceDate}`);
        } else {
          const jBody = await jobsResp.text();
          console.warn(`[ST customers] jobs lookup failed status=${jobsResp.status} body=${jBody.slice(0,200)}`);
        }
      } catch (je) {
        console.warn(`[ST customers] jobs lookup threw: ${je.message}`);
      }
    }

    res.json({
      customers: customers.map(c => ({
        id: c.id,
        name: c.name,
        type: c.type,
        active: c.active,
        balance: c.balance,
        doNotService: c.doNotService,
        contacts: (c.contacts || []).slice(0, 3).map(ct => ({
          name: ct.name,
          phones: (ct.phones || []).slice(0, 2).map(p => ({ type: p.type, value: p.value })),
          emails: (ct.emails || []).slice(0, 2).map(e => e.address),
        })),
        address: (() => {
          const a = (c.addresses || []).find(x => x.type === 'ServiceLocation') || c.addresses?.[0];
          return a ? [a.street, a.city, a.state, a.zip].filter(Boolean).join(', ') : null;
        })(),
      })),
      lastServiceDate,
    });
  } catch (e) {
    console.error(`[ST customers] 500 error: ${e.message}`, e.stack);
    res.status(500).json({ error: e.message, customers: [] });
  }
});

// ST pricebook proxy
app.post('/api/st', async (req, res) => {
  try {
    const data = await retryFetch(async () => {
      const response = await fetch(PROXY_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Accept-Encoding': 'identity' },
        body: JSON.stringify(req.body)
      });
      return response.json();
    });
    res.json(data);
  } catch (e) {
    console.error('st proxy error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Inventory lookup from cache
app.get('/api/inventory', (req, res) => {
  const term = (req.query.q || '').trim();
  const meta = { cacheAge: lastCacheUpdate, totalCached: inventoryCache.length, fromFallback: inventoryFromFallback, fallbackDate: fallbackCacheDate };
  if (!term) {
    return res.json({ items: [], ...meta });
  }
  const matches = searchInventoryCache(term);
  const byCode = {};
  matches.forEach(row => {
    if (!byCode[row.code]) byCode[row.code] = { code: row.code, name: row.name, rows: [] };
    byCode[row.code].rows.push(row);
  });

  const results = Object.values(byCode).map(item => ({
    code:      item.code,
    name:      item.name,
    inventory: aggregateInventory(item.rows),
    rawRows:   item.rows,
  }));

  res.json({ items: results, ...meta });
});

// Force inventory refresh from Gmail
app.post('/api/inventory/refresh', async (req, res) => {
  try {
    await fetchInventoryFromGmail();
    res.json({ success: true, totalRows: inventoryCache.length, updatedAt: lastCacheUpdate });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Cache status
app.get('/api/inventory/status', (req, res) => {
  res.json({ totalRows: inventoryCache.length, lastUpdated: lastCacheUpdate, isFetching });
});

// Estimates library
app.get('/api/estimates', (req, res) => {
  const list = readEstimates();
  res.json(list.slice().reverse()); // newest first
});

app.post('/api/estimates', async (req, res) => {
  const list = readEstimates();
  const record = { id: Date.now().toString(), timestamp: new Date().toISOString(), ...req.body };
  list.push(record);
  writeEstimates(list);
  if (req.body.type === 'site-survey') {
    appendToSheet(buildSheetRow(req.body)); // fire-and-forget; never blocks response
  }
  res.json(record);
});

app.get('/api/stories', (req, res) => {
  const list = readStories();
  res.json({ stories: list.slice().reverse() }); // newest first
});

app.post('/api/stories', (req, res) => {
  const list = readStories();
  const story = { id: Date.now().toString(), created_at: new Date().toISOString(), ...req.body };
  list.push(story);
  writeStories(list);
  res.json({ success: true, story });
});

app.get('/api/sheets/setup', async (req, res) => {
  try {
    const token = await getGmailToken();
    const range = encodeURIComponent(`${SHEET_NAME}!A1`);
    const r = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${range}?valueInputOption=USER_ENTERED`,
      {
        method: 'PUT',
        headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json', 'Accept-Encoding': 'identity' },
        body: JSON.stringify({ values: [SHEET_HEADERS] }),
      }
    );
    const data = await r.json();
    if (!r.ok) return res.status(500).json({ error: data });
    res.json({ ok: true, updatedRange: data.updatedRange });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.delete('/api/estimates/:id', (req, res) => {
  const list = readEstimates();
  const filtered = list.filter(e => e.id !== req.params.id);
  writeEstimates(filtered);
  res.json({ success: true });
});

// ─── MARKET INTEL ─────────────────────────────────────────────────────────────
app.post('/api/market-intel/compare', async (req, res) => {
  console.log('Compare request received');
  try {
    const { hodgeBase64, hodgeMimeType, compBase64, compMimeType } = req.body;
    if (!hodgeBase64 || !hodgeMimeType || !compBase64 || !compMimeType) {
      return res.status(400).json({ error: 'Missing file data' });
    }
    function makeBlock(base64, mimeType) {
      return mimeType.startsWith('image/')
        ? { type: 'image',    source: { type: 'base64', media_type: mimeType,          data: base64 } }
        : { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: base64 } };
    }
    const systemPrompt = `You are a precise market price analyst for an industrial compressed air company. You will receive two service quotes. Analyze them carefully, then output a single valid JSON object and nothing else. No markdown, no code fences, no commentary outside the JSON.`;
    const userMsg = `Compare these two quotes line by line. Group by equipment unit. Match line items by what they ARE (air filter = air filter regardless of part number). For units appearing in both quotes, match their line items side by side. For units in only one quote, include them with 0 for the missing side. Classify each line item as: Parts = filters/oil/separators/lubricant/consumables; Labor = labor hours/rates/travel/mileage; Equipment = compressors/dryers/tanks/major equipment being sold; Other = surcharges/fees/misc. Return this exact JSON structure — no extra fields, no duplicate keys, strictly valid JSON:\n\n{"competitor": "string", "customer": "string", "groups": [{"equipmentModel": "string", "lineItems": [{"description": "string", "category": "Parts|Labor|Equipment|Other", "hodgePrice": number, "competitorPrice": number}]}]}`;
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept-Encoding': 'identity',
        'x-api-key': process.env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
        'anthropic-beta': 'pdfs-2024-09-25',
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 4096,
        system: systemPrompt,
        messages: [{
          role: 'user',
          content: [
            { type: 'text', text: 'This is the HODGE quote:' },
            makeBlock(hodgeBase64, hodgeMimeType),
            { type: 'text', text: 'This is the COMPETITOR quote:' },
            makeBlock(compBase64, compMimeType),
            { type: 'text', text: userMsg },
          ],
        }],
      }),
    });
    const data = await response.json();
    if (data.error) throw new Error(data.error.message || 'Claude API error');
    const text = data.content[0].text.trim();
    console.log('Compare raw response received, length:', text.length, 'chars');
    const jsonStr = text.replace(/^```(?:json)?\s*\n?/, '').replace(/\n?```\s*$/, '').trim();
    let result;
    try {
      result = JSON.parse(jsonStr);
    } catch (parseErr) {
      console.error('JSON parse error:', parseErr.message, '\nRaw string was:', jsonStr);
      return res.status(422).json({ error: 'Claude returned malformed JSON: ' + parseErr.message });
    }
    console.log('Parsed top-level keys:', Object.keys(result));
    console.log('Parsed result groups:', result.groups ? result.groups.length : 0);
    res.json({ success: true, result });
  } catch (e) {
    console.error('market-intel/compare error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/market-intel/log', async (req, res) => {
  try {
    const { branch, repName, customer, competitor, equipmentModel, quoteDate, lineItems } = req.body;
    console.log('Log request received — branch:', branch, '| repName:', repName, '| items:', lineItems ? lineItems.length : 0);
    if (!lineItems || lineItems.length === 0) {
      return res.status(400).json({ error: 'No line items provided' });
    }
    const keyJson = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY);
    const auth = new google.auth.GoogleAuth({
      credentials: keyJson,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const sheets = google.sheets({ version: 'v4', auth });
    const timestamp = new Date().toISOString();
    const rows = lineItems.map(item => {
      const hodge = parseFloat(item.hodgePrice) || 0;
      const comp  = parseFloat(item.competitorPrice) || 0;
      const delta = hodge - comp;
      const pctDelta = comp === 0 ? 'N/A' : ((hodge - comp) / comp * 100).toFixed(2);
      return [
        timestamp,
        branch || '',
        repName || '',
        customer || '',
        competitor || '',
        item.equipmentModel || equipmentModel || '',
        item.description || '',
        item.category || '',
        hodge.toFixed(2),
        comp.toFixed(2),
        delta.toFixed(2),
        pctDelta,
        quoteDate || '',
      ];
    });
    await sheets.spreadsheets.values.append({
      spreadsheetId: '1avLUXAsxnDWu8h5qY1_BAMlmJqS8q183qeszw0Wynxk',
      range: 'Sheet1!A:M',
      valueInputOption: 'RAW',
      requestBody: { values: rows },
    });
    res.json({ success: true, rowsAdded: rows.length });
  } catch (e) {
    console.error('market-intel/log error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── HEALTH CHECK SYSTEM ─────────────────────────────────────────────────────
let healthCache = null;
let prevHealthStatus = {};

async function sendSlackAlert(text) {
  const webhookUrl = process.env.SLACK_STATUS_WEBHOOK_URL;
  if (!webhookUrl) return;
  try {
    await fetch(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Accept-Encoding': 'identity' },
      body: JSON.stringify({ text }),
    });
  } catch (e) { console.error('Slack alert failed:', e.message); }
}

async function checkAnthropic() {
  const t0 = Date.now();
  try {
    const resp = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Accept-Encoding': 'identity',
        'x-api-key': process.env.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({ model: 'claude-haiku-4-5-20251001', max_tokens: 1, messages: [{ role: 'user', content: 'hi' }] }),
    });
    return { status: resp.status < 500 ? 'up' : 'down', latencyMs: Date.now() - t0 };
  } catch (e) { return { status: 'down', latencyMs: Date.now() - t0, error: e.message }; }
}

async function checkGmail() {
  const t0 = Date.now();
  try {
    const oauth2Client = new google.auth.OAuth2(process.env.GMAIL_CLIENT_ID, process.env.GMAIL_CLIENT_SECRET);
    oauth2Client.setCredentials({ refresh_token: process.env.GMAIL_REFRESH_TOKEN });
    const gmail = google.gmail({ version: 'v1', auth: oauth2Client });
    await gmail.users.getProfile({ userId: 'me' });
    return { status: 'up', latencyMs: Date.now() - t0 };
  } catch (e) { return { status: 'down', latencyMs: Date.now() - t0, error: e.message }; }
}

async function checkMonday() {
  if (!process.env.MONDAY_API_KEY) return { status: 'down', latencyMs: 0, error: 'Not configured' };
  const t0 = Date.now();
  try {
    const data = await mondayGraphQL(`query { me { id } }`);
    return { status: data.data?.me?.id ? 'up' : 'down', latencyMs: Date.now() - t0 };
  } catch (e) { return { status: 'down', latencyMs: Date.now() - t0, error: e.message }; }
}

async function checkServiceTitan() {
  const t0 = Date.now();
  try {
    await fetch(PROXY_URL, {
      method: 'GET',
      headers: { 'Accept-Encoding': 'identity' },
    });
    return { status: 'up', latencyMs: Date.now() - t0 }; // any HTTP response = worker is reachable
  } catch (e) { return { status: 'down', latencyMs: Date.now() - t0, error: e.message }; }
}

async function checkPostgres() {
  const t0 = Date.now();
  if (!pgPool) return { status: 'down', latencyMs: 0, error: 'Not configured' };
  try {
    await pgPool.query('SELECT 1');
    return { status: 'up', latencyMs: Date.now() - t0 };
  } catch (e) { return { status: 'down', latencyMs: Date.now() - t0, error: e.message }; }
}

async function runHealthCheck() {
  const results = await Promise.allSettled([
    checkAnthropic(), checkGmail(), checkMonday(), checkServiceTitan(), checkPostgres(),
  ]);
  const [anthropic, gmail, monday, servicetitan, postgres] = results.map(r =>
    r.status === 'fulfilled' ? r.value : { status: 'down', latencyMs: 0, error: r.reason?.message }
  );
  const services = { anthropic, gmail, monday, servicetitan, postgres };
  const downCount = Object.values(services).filter(s => s.status === 'down').length;
  const overallStatus = downCount === 0 ? 'healthy' : downCount >= 3 ? 'down' : 'degraded';

  for (const [name, svc] of Object.entries(services)) {
    const prev = prevHealthStatus[name];
    if (prev && prev !== svc.status) {
      if (svc.status === 'down') sendSlackAlert(`🔴 *AirIQ Alert:* ${name} is DOWN`);
      else if (prev === 'down') sendSlackAlert(`🟢 *AirIQ Recovery:* ${name} is back UP`);
    }
    prevHealthStatus[name] = svc.status;
  }

  healthCache = { status: overallStatus, services, checkedAt: new Date().toISOString() };
  console.log(`Health check: ${overallStatus} — ${JSON.stringify(Object.fromEntries(Object.entries(services).map(([k,v])=>[k,v.status])))}`);
  return healthCache;
}

app.get('/api/health', (req, res) => {
  if (!healthCache) return res.json({ status: 'starting', checkedAt: null });
  res.json(healthCache);
});

app.post('/api/health/test-alert', async (req, res) => {
  await sendSlackAlert('🧪 *AirIQ Status Test* — Slack alert webhook is working ✅');
  res.json({ ok: true, message: 'Test alert sent to SLACK_STATUS_WEBHOOK_URL' });
});

app.get('/status', (req, res) => {
  const h = healthCache;
  if (!h) return res.send(`<!DOCTYPE html><html><head><meta charset="utf-8"><meta http-equiv="refresh" content="5"><title>AirIQ Status</title><style>body{background:#0a1628;color:#fff;font-family:sans-serif;padding:40px}</style></head><body><h1 style="color:#f97316">AirIQ Status</h1><p>Health check not yet run — will refresh in 5 seconds.</p></body></html>`);

  const dot = s => { if (!s) return '🔴'; if (s.status === 'up') return s.latencyMs > 2000 ? '🟡' : '🟢'; return '🔴'; };
  const rows = Object.entries(h.services).map(([name, s]) =>
    `<tr><td>${dot(s)}&nbsp;&nbsp;${name.charAt(0).toUpperCase()+name.slice(1)}</td>` +
    `<td style="color:${s.status==='up'?'#4ade80':'#f87171'};font-weight:bold">${s.status.toUpperCase()}</td>` +
    `<td>${s.latencyMs != null ? s.latencyMs+'&nbsp;ms' : '—'}</td>` +
    `<td style="color:#94a3b8;font-size:12px">${s.error||''}</td></tr>`
  ).join('');
  const statusColor = h.status === 'healthy' ? '#4ade80' : h.status === 'degraded' ? '#fbbf24' : '#f87171';
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.send(`<!DOCTYPE html>
<html><head>
  <meta charset="utf-8">
  <meta http-equiv="refresh" content="30">
  <title>AirIQ Status</title>
  <style>
    *{box-sizing:border-box}
    body{background:#0a1628;color:#e2e8f0;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;padding:40px;margin:0}
    h1{color:#f97316;margin:0 0 8px;font-size:26px;letter-spacing:-.02em}
    .badge{display:inline-block;padding:4px 16px;border-radius:20px;font-weight:700;font-size:13px;
      background:${statusColor}1a;color:${statusColor};border:1px solid ${statusColor}60;margin-bottom:8px}
    .ts{color:#475569;font-size:13px;margin-bottom:28px}
    table{border-collapse:collapse;width:100%;max-width:620px}
    th{text-align:left;color:#f97316;border-bottom:2px solid #1e3a5f;padding:8px 16px;font-size:12px;text-transform:uppercase;letter-spacing:.08em}
    td{padding:13px 16px;border-bottom:1px solid #1e3a5f;font-size:15px;vertical-align:middle}
    tr:last-child td{border-bottom:none}
    .note{color:#334155;font-size:12px;margin-top:20px}
  </style>
</head><body>
  <h1>AirIQ Status</h1>
  <div class="badge">${h.status.toUpperCase()}</div>
  <div class="ts">Last checked: ${new Date(h.checkedAt).toLocaleString('en-US',{timeZone:'America/New_York'})} EST</div>
  <table>
    <thead><tr><th>Service</th><th>Status</th><th>Latency</th><th>Notes</th></tr></thead>
    <tbody>${rows}</tbody>
  </table>
  <div class="note">Auto-refreshes every 30 seconds</div>
</body></html>`);
});

// ─── SERIAL LOOKUP ───────────────────────────────────────────────────────────
const SERIAL_SHEET_ID  = '15JjnuJuxvJpnj218Lr0s_xTwZazOLTrr2h5DXNJ9nVo';
const SERIAL_SHEET_TAB = 'Hodge Compressor Master Sales';
let serialCache = { rows: null, fetchedAt: 0 };

async function fetchSerialSheet() {
  if (serialCache.rows && (Date.now() - serialCache.fetchedAt) < 5 * 60 * 1000) {
    return serialCache.rows;
  }
  const keyJson = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY);
  const auth = new google.auth.GoogleAuth({
    credentials: keyJson,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SERIAL_SHEET_ID,
    range: SERIAL_SHEET_TAB,
  });
  const rows = res.data.values || [];
  serialCache = { rows, fetchedAt: Date.now() };
  console.log(`Serial sheet cached: ${rows.length} rows`);
  return rows;
}

app.get('/api/serial-lookup', async (req, res) => {
  const serial = (req.query.serial || '').trim();
  if (!serial) return res.status(400).json({ error: 'serial parameter required' });
  try {
    const rows = await fetchSerialSheet();
    if (rows.length === 0) return res.json({ matches: [], total: 0 });
    const headers = rows[0].map(h => (h || '').toString().trim().toLowerCase());
    const ci = kw => headers.findIndex(h => h.includes(kw));
    const serialIdx  = ci('serial');
    const modelIdx   = ci('model');
    const custIdx    = ci('customer');
    const saleDateIdx = (() => { const i = ci('sale date'); return i >= 0 ? i : ci('sale'); })();
    const startupIdx = ci('startup');
    const memoIdx    = ci('memo');
    if (serialIdx === -1) {
      console.error('Serial column not found. Headers:', headers);
      return res.status(500).json({ error: 'Serial number column not found in sheet. Headers: ' + headers.join(', ') });
    }
    const q = serial.toLowerCase();
    const matches = rows.slice(1)
      .filter(row => (row[serialIdx] || '').toString().toLowerCase().includes(q))
      .map(row => ({
        serialNumber: row[serialIdx]   || '',
        model:        modelIdx   >= 0 ? (row[modelIdx]    || '') : '',
        customer:     custIdx    >= 0 ? (row[custIdx]     || '') : '',
        saleDate:     saleDateIdx >= 0 ? (row[saleDateIdx] || '') : '',
        startupDate:  startupIdx >= 0 ? (row[startupIdx]  || '') : '',
        memo:         memoIdx    >= 0 ? (row[memoIdx]     || '') : '',
      }));
    res.json({ matches, total: matches.length });
  } catch (e) {
    console.error('serial-lookup error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── FAQ ──────────────────────────────────────────────────────────────────────
const FAQ_SHEET_ID  = '1SwIQrNrBRosTdD-Vb_5ZsNaik5ChDMzhkpeJQ2PKuU4';
const FAQ_SHEET_TAB = 'Sheet1';
let faqCache = { items: null, fetchedAt: 0 };

async function fetchFaqSheet() {
  if (faqCache.items && (Date.now() - faqCache.fetchedAt) < 60 * 1000) {
    return faqCache.items;
  }
  const keyJson = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY);
  const auth = new google.auth.GoogleAuth({
    credentials: keyJson,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
  const sheets = google.sheets({ version: 'v4', auth });
  let rows = [];
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: FAQ_SHEET_ID,
      range: FAQ_SHEET_TAB,
    });
    rows = res.data.values || [];
  } catch (e) {
    if (e.message && (e.message.includes('Unable to parse range') || e.message.includes('notFound'))) {
      console.log('FAQ sheet tab not found yet — returning empty');
    } else {
      throw e;
    }
  }
  if (rows.length < 2) {
    faqCache = { items: [], fetchedAt: Date.now() };
    return [];
  }
  const headers = rows[0].map(h => (h || '').toString().trim().toLowerCase());
  const qi   = headers.findIndex(h => h.includes('question'));
  const ai   = headers.findIndex(h => h.includes('answer'));
  const ci   = headers.findIndex(h => h.includes('category'));
  const acti = headers.findIndex(h => h.includes('active'));
  const items = rows.slice(1)
    .filter(row => acti < 0 || (row[acti] || '').toString().trim().toLowerCase() === 'yes')
    .map(row => ({
      question: qi >= 0 ? (row[qi] || '').toString().trim() : '',
      answer:   ai >= 0 ? (row[ai] || '').toString().trim() : '',
      category: ci >= 0 ? (row[ci] || '').toString().trim() || 'General' : 'General',
    }))
    .filter(item => item.question);
  faqCache = { items, fetchedAt: Date.now() };
  console.log(`FAQ cache loaded: ${items.length} active items`);
  return items;
}

app.get('/api/faq', async (req, res) => {
  try {
    const items = await fetchFaqSheet();
    res.json(items);
  } catch (e) {
    console.error('faq error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── MONDAY.COM PROSPECTS INTEGRATION ────────────────────────────────────────
const MONDAY_BOARD_ID = 18418669668;
const REP_NAME_MAP = {
  'Devon':       'Devon Dalton',
  'Earl':        'Earl Bryan Livingston',
  'Jack':        'Jack Wolchek',
  'Connor':      'Connor Smith',
  'Dan':         'Dan Constant',
  'Dave J':      'Dave Jolly',
  'Jon L':       'Jon Landress',
  'Chase':       'Chase Lewis',
  'Zac':         'Zac Hoadley',
  'Mike':        'Mike George',
  'Jackson':     'Jackson Carrell',
  'Tristian':    'Tristian Castiglione',
  'David G':     'David Gonzalez',
  'Javier':      'Javier Cabral',
  'Jon S':       'Jon Sanders',
  'Tony':        'Tony G',
  'Kyle':        'Kyle White',
};
let mondayColMap  = null;
let mondayGroupId = null;
let mondayUserMap = {};
let mondayColTypes = {};

async function mondayGraphQL(query, variables) {
  const body = variables ? { query, variables } : { query };
  const data = await retryFetch(async () => {
    const res = await fetch('https://api.monday.com/v2', {
      method: 'POST',
      headers: {
        'Content-Type':    'application/json',
        'Authorization':   `Bearer ${process.env.MONDAY_API_KEY}`,
        'Accept-Encoding': 'identity',
      },
      body: JSON.stringify(body),
    });
    return res.json();
  });
  if (data.errors) throw new Error(data.errors.map(e => e.message).join('; '));
  return data;
}

async function fetchMondayMeta() {
  const data = await mondayGraphQL(`
    query {
      boards(ids: [${MONDAY_BOARD_ID}]) {
        columns { id title type }
        groups  { id title }
      }
    }
  `);
  const board = data.data?.boards?.[0];
  if (!board) throw new Error('Monday.com board not found');

  const byTitle = {};
  const colTypeMap = {};
  board.columns.forEach(c => { byTitle[c.title] = c.id; colTypeMap[c.id] = c.type; });
  mondayColTypes = colTypeMap;

  mondayColMap = {
    rep:         byTitle['Rep']          || null,
    status:      byTitle['Status']       || null,
    lastVisited: byTitle['Last Visited'] || null,
    contactName: byTitle['Contact Name'] || null,
    branch:      byTitle['Branch']       || null,
    outcome:     byTitle['Outcome']      || null,
    location:    byTitle['Location']     || null,
    notes:       byTitle['Notes']        || null,
    phone:       byTitle['Phone']        || null,
    email:       byTitle['Email']        || null,
    stopCount:   byTitle['Stop Count']   || null,
  };

  const activeGroup = board.groups.find(g => g.title.toLowerCase().includes('active prospect'));
  mondayGroupId = activeGroup?.id || null;
  console.log('Monday.com metadata cached — columns:', JSON.stringify(mondayColMap), '| group:', mondayGroupId);
}

async function fetchMondayUsers() {
  const data = await mondayGraphQL(`query { users { id name } }`);
  const users = data.data?.users || [];
  mondayUserMap = {};
  users.forEach(u => { mondayUserMap[u.name.toLowerCase()] = Number(u.id); });
  console.log(`Monday.com users cached: ${users.length} users`);
}

async function mondayUpsertProspect(stopData) {
  try {
    if (!process.env.MONDAY_API_KEY) return;
    if (!mondayColMap) await fetchMondayMeta();

    const { repName, companyName, contactName, location, notableMoment,
            pmOpp, equipOpp, serviceLead, pipingOpp, sticker, vrPres,
            apptSet, nothing, stickerCount, mobile, officePhone, email, branch } = stopData;

    if (!companyName) return;

    const outcomes = [];
    if (pmOpp)       outcomes.push('Membership Opp');
    if (equipOpp)    outcomes.push('Equipment Opp');
    if (serviceLead) outcomes.push('Service Lead');
    if (pipingOpp)   outcomes.push('Piping/Install');
    if (sticker)     outcomes.push(Number(stickerCount) > 1 ? `Sticker Placed x${stickerCount}` : 'Sticker Placed');
    if (vrPres)      outcomes.push('VR Presentation');
    if (apptSet)     outcomes.push('Appointment Set');
    if (nothing)     outcomes.push('Nothing');
    const outcomeStr = outcomes.join(', ');

    const today = new Date().toISOString().slice(0, 10);
    const mondayDisplayName = REP_NAME_MAP[repName] || repName;
    const mondayUserId = mondayUserMap[mondayDisplayName.toLowerCase()] || null;

    // Search for existing item by exact company name (inlined — monday.com v2 requires CompareValue! scalar, not [String])
    const safeCompanyName = companyName.replace(/"/g, '\\"');
    const searchQuery = `
      query {
        boards(ids: [${MONDAY_BOARD_ID}]) {
          items_page(limit: 5, query_params: {
            rules: [{ column_id: "name", compare_value: "${safeCompanyName}" }]
          }) {
            items {
              id
              name
              column_values { id text }
            }
          }
        }
      }
    `;
    const searchData = await retryFetch(async () => {
      const searchRaw = await fetch('https://api.monday.com/v2', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${process.env.MONDAY_API_KEY}`,
          'Content-Type': 'application/json',
          'API-Version': '2024-01',
        },
        body: JSON.stringify({ query: searchQuery }),
      });
      return searchRaw.json();
    });
    if (searchData.errors) throw new Error(searchData.errors.map(e => e.message).join('; '));

    const items = searchData.data?.boards?.[0]?.items_page?.items || [];

    // Exact name + rep match — compare against monday.com display name (People column returns display name as text)
    const match = items.find(item => {
      if (item.name.toLowerCase() !== companyName.toLowerCase()) return false;
      if (!mondayColMap.rep) return true;
      const repCol = item.column_values.find(cv => cv.id === mondayColMap.rep);
      return repCol && repCol.text.toLowerCase() === mondayDisplayName.toLowerCase();
    });

    if (match) {
      const scCol = mondayColMap.stopCount ? match.column_values.find(cv => cv.id === mondayColMap.stopCount) : null;
      const currentCount = scCol ? (parseInt(scCol.text, 10) || 0) : 0;

      const colValues = {};
      if (mondayColMap.lastVisited)                  colValues[mondayColMap.lastVisited] = { date: today };
      if (mondayColMap.outcome     && outcomeStr)    colValues[mondayColMap.outcome]     = outcomeStr;
      if (mondayColMap.notes       && notableMoment) colValues[mondayColMap.notes]       = notableMoment;
      if (mondayColMap.contactName && contactName)   colValues[mondayColMap.contactName] = contactName;
      const phoneVal = mobile || officePhone;
      if (mondayColMap.phone && phoneVal)
        colValues[mondayColMap.phone] = mondayColTypes[mondayColMap.phone] === 'phone'
          ? { phone: phoneVal, countryShortName: 'US' } : phoneVal;
      if (mondayColMap.email && email)
        colValues[mondayColMap.email] = mondayColTypes[mondayColMap.email] === 'email'
          ? { email, text: email } : email;
      if (mondayColMap.location    && location)      colValues[mondayColMap.location]    = location;
      if (mondayColMap.branch      && branch)        colValues[mondayColMap.branch]       = { labels: [branch] };
      if (mondayColMap.stopCount)                    colValues[mondayColMap.stopCount]   = currentCount + 1;

      await mondayGraphQL(`
        mutation {
          change_multiple_column_values(
            board_id: ${MONDAY_BOARD_ID},
            item_id: ${match.id},
            column_values: ${JSON.stringify(JSON.stringify(colValues))}
          ) { id }
        }
      `);
      console.log(`Monday.com: updated "${companyName}" (${repName}) -- stops: ${currentCount + 1}`);
    } else {
      const colValues = {};
      if (mondayColMap.rep) {
        const repIsText = mondayColTypes[mondayColMap.rep] === 'text' || mondayColMap.rep.startsWith('text_');
        if (repIsText) {
          colValues[mondayColMap.rep] = mondayDisplayName;
        } else if (mondayUserId) {
          colValues[mondayColMap.rep] = { personsAndTeams: [{ id: mondayUserId, kind: 'person' }] };
        }
      }
      if (mondayColMap.status)                       colValues[mondayColMap.status]      = { label: 'Active Prospects' };
      if (mondayColMap.lastVisited)                  colValues[mondayColMap.lastVisited] = { date: today };
      if (mondayColMap.outcome     && outcomeStr)    colValues[mondayColMap.outcome]     = outcomeStr;
      if (mondayColMap.notes       && notableMoment) colValues[mondayColMap.notes]       = notableMoment;
      if (mondayColMap.contactName && contactName)   colValues[mondayColMap.contactName] = contactName;
      const phoneVal = mobile || officePhone;
      if (mondayColMap.phone && phoneVal)
        colValues[mondayColMap.phone] = mondayColTypes[mondayColMap.phone] === 'phone'
          ? { phone: phoneVal, countryShortName: 'US' } : phoneVal;
      if (mondayColMap.email && email)
        colValues[mondayColMap.email] = mondayColTypes[mondayColMap.email] === 'email'
          ? { email, text: email } : email;
      if (mondayColMap.location    && location)      colValues[mondayColMap.location]    = location;
      if (mondayColMap.branch      && branch)        colValues[mondayColMap.branch]       = { labels: [branch] };
      if (mondayColMap.stopCount)                    colValues[mondayColMap.stopCount]   = 1;

      const groupClause = mondayGroupId ? `group_id: "${mondayGroupId}",` : '';
      await mondayGraphQL(`
        mutation {
          create_item(
            board_id: ${MONDAY_BOARD_ID},
            ${groupClause}
            item_name: ${JSON.stringify(companyName)},
            column_values: ${JSON.stringify(JSON.stringify(colValues))}
          ) { id }
        }
      `);
      console.log(`Monday.com: created prospect "${companyName}" (${repName})`);
    }
  } catch (e) {
    console.error('mondayUpsertProspect error:', e.message);
  }
}

// ─── FIELD LOG ────────────────────────────────────────────────────────────────
app.post('/api/field-log', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  try {
    const { repName, pmOpp, equipOpp, serviceLead, pipingOpp, sticker, vrPres, apptSet, nothing,
            location, notableMoment, companyName, contactName, contactTitle, stickerCount,
            mobile, officePhone, email, website, cardAddress, branch } = req.body;
    if (!repName) return res.status(400).json({ error: 'repName required' });
    const sc = sticker ? Math.max(1, Number(stickerCount) || 1) : 0;
    const { rows } = await pgPool.query(
      `INSERT INTO field_log_entries
         (rep_name, pm_opp, equip_opp, service_lead, piping_opp, sticker, vr_pres, appt_set, nothing,
          location, notable_moment, company_name, contact_name, contact_title, sticker_count,
          mobile, office_phone, email, website, card_address)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20)
       RETURNING id, logged_at`,
      [repName, !!pmOpp, !!equipOpp, !!serviceLead, !!pipingOpp, !!sticker, !!vrPres, !!apptSet, !!nothing,
       location || null, notableMoment || null, companyName || null, contactName || null, contactTitle || null, sc,
       mobile || null, officePhone || null, email || null, website || null, cardAddress || null]
    );
    res.json({ success: true, id: rows[0].id, loggedAt: rows[0].logged_at });
    if (companyName && process.env.MONDAY_API_KEY) {
      mondayUpsertProspect({
        repName, companyName, contactName, location, notableMoment,
        pmOpp: !!pmOpp, equipOpp: !!equipOpp, serviceLead: !!serviceLead,
        pipingOpp: !!pipingOpp, sticker: !!sticker, vrPres: !!vrPres,
        apptSet: !!apptSet, nothing: !!nothing,
        stickerCount: sc, mobile, officePhone, email, branch: branch || null,
      }).catch(e => console.error('Monday upsert error:', e.message));
    }
  } catch (e) {
    console.error('field-log POST error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/field-log/today/:repName', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  try {
    const TZ = getTZ(req.query.branch);
    const { rows } = await pgPool.query(
      `SELECT
         COUNT(*)::int               AS stops,
         SUM(pm_opp::int)::int       AS pm_opps,
         SUM(equip_opp::int)::int    AS equip_opps,
         SUM(service_lead::int)::int AS service_leads,
         SUM(piping_opp::int)::int   AS piping_opps,
         SUM(COALESCE(sticker_count, sticker::int))::int AS stickers,
         SUM(vr_pres::int)::int      AS vr_pres,
         SUM(appt_set::int)::int     AS appt_set,
         SUM(nothing::int)::int      AS nothing
       FROM field_log_entries
       WHERE rep_name = $1
         AND logged_at >= DATE_TRUNC('day', NOW() AT TIME ZONE '${TZ}') AT TIME ZONE '${TZ}'`,
      [req.params.repName]
    );
    res.json(rows[0]);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/field-log/dashboard', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  try {
    const range = req.query.range || 'today';
    const TZ = 'America/New_York';
    // Use Eastern-time date truncation so midnight is midnight in EST, not UTC
    let whereClause;
    if (range === 'week') {
      whereClause = `WHERE logged_at >= DATE_TRUNC('week', NOW() AT TIME ZONE '${TZ}') AT TIME ZONE '${TZ}'`;
    } else if (range === 'month') {
      whereClause = `WHERE logged_at >= DATE_TRUNC('month', NOW() AT TIME ZONE '${TZ}') AT TIME ZONE '${TZ}'`;
    } else if (range === 'all') {
      whereClause = '';
    } else {
      whereClause = `WHERE logged_at >= DATE_TRUNC('day', NOW() AT TIME ZONE '${TZ}') AT TIME ZONE '${TZ}'`;
    }
    const weekStopsExpr = range === 'today'
      ? `(SELECT COUNT(*)::int FROM field_log_entries w WHERE w.rep_name = fe.rep_name AND w.logged_at >= DATE_TRUNC('week', NOW() AT TIME ZONE '${TZ}') AT TIME ZONE '${TZ}')`
      : `0`;
    const { rows } = await pgPool.query(
      `SELECT
         rep_name,
         COUNT(*)::int               AS stops,
         SUM(pm_opp::int)::int       AS pm_opps,
         SUM(equip_opp::int)::int    AS equip_opps,
         SUM(service_lead::int)::int AS service_leads,
         SUM(piping_opp::int)::int   AS piping_opps,
         SUM(COALESCE(sticker_count, sticker::int))::int AS stickers,
         SUM(vr_pres::int)::int      AS vr_pres,
         SUM(appt_set::int)::int     AS appt_set,
         SUM(nothing::int)::int      AS nothing,
         MAX(logged_at)              AS last_logged,
         SUM(CASE WHEN (pm_opp OR equip_opp OR service_lead OR piping_opp OR appt_set OR vr_pres) THEN 1 ELSE 0 END)::int AS opp_stops,
         ROUND(
           CASE WHEN COUNT(*) = 0 THEN 0::numeric
           ELSE SUM(CASE WHEN (pm_opp OR equip_opp OR service_lead OR piping_opp OR appt_set OR vr_pres) THEN 1 ELSE 0 END)::numeric / COUNT(*) * 100
           END, 1
         ) AS conversion_rate,
         ${weekStopsExpr}::int AS week_stops
       FROM field_log_entries fe
       ${whereClause}
       GROUP BY rep_name
       ORDER BY stops DESC`
    );
    res.json(rows);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/field-log/week', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  try {
    const { rows } = await pgPool.query(
      `SELECT
         (logged_at AT TIME ZONE 'America/New_York')::date::text AS day,
         COUNT(*)::int AS stops
       FROM field_log_entries
       WHERE logged_at >= DATE_TRUNC('week', NOW() AT TIME ZONE 'America/New_York') AT TIME ZONE 'America/New_York'
       GROUP BY 1
       ORDER BY 1`
    );
    res.json(rows);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/field-log/trend', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  const range = req.query.range || 'week';
  const TZ = 'America/New_York';
  try {
    let query, bucket;
    if (range === 'today') {
      query = `SELECT
         EXTRACT(HOUR FROM logged_at AT TIME ZONE '${TZ}')::int AS hour,
         COUNT(*)::int AS stops
       FROM field_log_entries
       WHERE logged_at >= DATE_TRUNC('day', NOW() AT TIME ZONE '${TZ}') AT TIME ZONE '${TZ}'
       GROUP BY 1 ORDER BY 1`;
      bucket = 'hour';
    } else if (range === 'week') {
      query = `SELECT
         (logged_at AT TIME ZONE '${TZ}')::date::text AS day,
         COUNT(*)::int AS stops
       FROM field_log_entries
       WHERE logged_at >= DATE_TRUNC('week', NOW() AT TIME ZONE '${TZ}') AT TIME ZONE '${TZ}'
       GROUP BY 1 ORDER BY 1`;
      bucket = 'day';
    } else if (range === 'month') {
      query = `SELECT
         (logged_at AT TIME ZONE '${TZ}')::date::text AS day,
         COUNT(*)::int AS stops
       FROM field_log_entries
       WHERE logged_at >= DATE_TRUNC('month', NOW() AT TIME ZONE '${TZ}') AT TIME ZONE '${TZ}'
       GROUP BY 1 ORDER BY 1`;
      bucket = 'day';
    } else {
      // all-time: weekly if data span < 6 months, monthly if longer
      const spanRes = await pgPool.query(
        `SELECT MIN(logged_at) AS min_t, MAX(logged_at) AS max_t FROM field_log_entries`
      );
      const { min_t, max_t } = spanRes.rows[0];
      const diffMonths = (min_t && max_t)
        ? (new Date(max_t) - new Date(min_t)) / (1000 * 60 * 60 * 24 * 30)
        : 0;
      if (diffMonths < 6) {
        query = `SELECT
           TO_CHAR(DATE_TRUNC('week', logged_at AT TIME ZONE '${TZ}'), 'YYYY-MM-DD') AS week,
           COUNT(*)::int AS stops
         FROM field_log_entries
         GROUP BY 1 ORDER BY 1`;
        bucket = 'week';
      } else {
        query = `SELECT
           TO_CHAR(DATE_TRUNC('month', logged_at AT TIME ZONE '${TZ}'), 'YYYY-MM') AS month,
           COUNT(*)::int AS stops
         FROM field_log_entries
         GROUP BY 1 ORDER BY 1`;
        bucket = 'month';
      }
    }
    const { rows } = await pgPool.query(query);
    res.json({ bucket, rows });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.patch('/api/field-log/stop/:id', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  const id = parseInt(req.params.id, 10);
  if (!id) return res.status(400).json({ error: 'Invalid stop id' });
  const {
    company_name, contact_name, location, notes,
    pm_opp, equip_opp, service_lead, piping_opp,
    sticker_placed, sticker_count, vr_pres, appt_set, nothing,
  } = req.body;
  try {
    const { rows } = await pgPool.query(
      `UPDATE field_log_entries SET
         company_name   = $1,
         contact_name   = $2,
         location       = $3,
         notable_moment = $4,
         pm_opp         = $5,
         equip_opp      = $6,
         service_lead   = $7,
         piping_opp     = $8,
         sticker        = $9,
         sticker_count  = $10,
         vr_pres        = $11,
         appt_set       = $12,
         nothing        = $13
       WHERE id = $14
       RETURNING id, rep_name, logged_at, company_name, contact_name, location, notable_moment,
                 pm_opp, equip_opp, service_lead, piping_opp, sticker, sticker_count,
                 vr_pres, appt_set, nothing, mobile, office_phone, email, website, card_address`,
      [
        company_name || null, contact_name || null, location || null, notes || null,
        !!pm_opp, !!equip_opp, !!service_lead, !!piping_opp,
        !!sticker_placed, sticker_placed ? Math.max(1, Number(sticker_count) || 1) : 0,
        !!vr_pres, !!appt_set, !!nothing,
        id,
      ]
    );
    if (rows.length === 0) return res.status(404).json({ error: 'Stop not found' });
    const updated = rows[0];
    res.json(updated);
    if (updated.company_name && process.env.MONDAY_API_KEY) {
      mondayUpsertProspect({
        repName:       updated.rep_name,
        companyName:   updated.company_name,
        contactName:   updated.contact_name,
        location:      updated.location,
        notableMoment: updated.notable_moment,
        pmOpp:         updated.pm_opp,
        equipOpp:      updated.equip_opp,
        serviceLead:   updated.service_lead,
        pipingOpp:     updated.piping_opp,
        sticker:       updated.sticker,
        vrPres:        updated.vr_pres,
        apptSet:       updated.appt_set,
        nothing:       updated.nothing,
        stickerCount:  updated.sticker_count,
        mobile:        updated.mobile,
        officePhone:   updated.office_phone,
        email:         updated.email,
        branch:        null,
      }).catch(e => console.error('Monday upsert error (PATCH):', e.message));
    }
  } catch (e) { res.status(500).json({ error: e.message }); }
});

async function mondayDecrementStop(companyName) {
  try {
    if (!mondayColMap) await fetchMondayMeta();
    if (!mondayColMap?.stopCount) return;
    const safe = companyName.replace(/"/g, '\\"');
    const searchData = await mondayGraphQL(`
      query {
        boards(ids: [${MONDAY_BOARD_ID}]) {
          items_page(limit: 5, query_params: {
            rules: [{ column_id: "name", compare_value: "${safe}" }]
          }) { items { id column_values { id text } } }
        }
      }
    `);
    const match = (searchData.data?.boards?.[0]?.items_page?.items || [])[0];
    if (!match) return;
    const scCol = match.column_values.find(cv => cv.id === mondayColMap.stopCount);
    const current = Number(scCol?.text) || 0;
    const next = Math.max(0, current - 1);
    const colValues = {};
    colValues[mondayColMap.stopCount] = next;
    await mondayGraphQL(`
      mutation {
        change_multiple_column_values(
          board_id: ${MONDAY_BOARD_ID},
          item_id: ${match.id},
          column_values: ${JSON.stringify(JSON.stringify(colValues))}
        ) { id }
      }
    `);
    console.log(`Monday.com: decremented stop count for "${companyName}" to ${next}`);
  } catch (e) {
    console.error('mondayDecrementStop error:', e.message);
  }
}

app.delete('/api/field-log/stop/:id', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  const id = parseInt(req.params.id, 10);
  if (!id) return res.status(400).json({ error: 'Invalid stop id' });
  try {
    const { rows } = await pgPool.query(
      `DELETE FROM field_log_entries WHERE id = $1 RETURNING rep_name, company_name`,
      [id]
    );
    if (rows.length === 0) return res.status(404).json({ error: 'Stop not found' });
    res.json({ success: true });
    const { company_name } = rows[0];
    if (company_name && process.env.MONDAY_API_KEY) {
      mondayDecrementStop(company_name).catch(() => {});
    }
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/field-log/all-stops', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  const range = req.query.range || 'today';
  const TZ = 'America/New_York';
  let whereClause;
  if (range === 'week') {
    whereClause = `WHERE logged_at >= DATE_TRUNC('week', NOW() AT TIME ZONE '${TZ}') AT TIME ZONE '${TZ}'`;
  } else if (range === 'month') {
    whereClause = `WHERE logged_at >= DATE_TRUNC('month', NOW() AT TIME ZONE '${TZ}') AT TIME ZONE '${TZ}'`;
  } else if (range === 'all') {
    whereClause = '';
  } else {
    whereClause = `WHERE logged_at >= DATE_TRUNC('day', NOW() AT TIME ZONE '${TZ}') AT TIME ZONE '${TZ}'`;
  }
  try {
    const { rows } = await pgPool.query(
      `SELECT id, rep_name, logged_at, company_name, contact_name, location, notable_moment,
              pm_opp, equip_opp, service_lead, piping_opp, sticker, sticker_count,
              vr_pres, appt_set, nothing, mobile, office_phone, email, website, card_address
       FROM field_log_entries
       ${whereClause}
       ORDER BY rep_name ASC, logged_at DESC
       LIMIT 2000`
    );
    res.json(rows);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/field-log/stops/:repName', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  try {
    const TZ = getTZ(req.query.branch);
    const { rows } = await pgPool.query(
      `SELECT id, logged_at, pm_opp, equip_opp, service_lead, piping_opp, sticker, vr_pres, appt_set, nothing,
              location, notable_moment, company_name, contact_name, sticker_count,
              mobile, office_phone, email, website, card_address
       FROM field_log_entries
       WHERE rep_name = $1
         AND logged_at >= DATE_TRUNC('day', NOW() AT TIME ZONE '${TZ}') AT TIME ZONE '${TZ}'
       ORDER BY logged_at DESC`,
      [req.params.repName]
    );
    res.json(rows);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/field-log/stops/:repName/all', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  try {
    const range = req.query.range || 'all';
    const TZ = getTZ(req.query.branch);
    let dateFilter = '';
    let limitClause = 'LIMIT 2000';
    if (range === 'today') {
      dateFilter = `AND logged_at >= DATE_TRUNC('day', NOW() AT TIME ZONE '${TZ}') AT TIME ZONE '${TZ}'`;
      limitClause = '';
    } else if (range === 'week') {
      dateFilter = `AND logged_at >= DATE_TRUNC('week', NOW() AT TIME ZONE '${TZ}') AT TIME ZONE '${TZ}'`;
      limitClause = '';
    } else if (range === 'month') {
      dateFilter = `AND logged_at >= DATE_TRUNC('month', NOW() AT TIME ZONE '${TZ}') AT TIME ZONE '${TZ}'`;
      limitClause = '';
    }
    const { rows } = await pgPool.query(
      `SELECT id, logged_at, pm_opp, equip_opp, service_lead, piping_opp, sticker, vr_pres, appt_set, nothing,
              location, notable_moment, company_name, contact_name, sticker_count,
              mobile, office_phone, email, website, card_address
       FROM field_log_entries
       WHERE rep_name = $1 ${dateFilter}
       ORDER BY logged_at DESC
       ${limitClause}`,
      [req.params.repName]
    );
    res.json(rows);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/field-log/export', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  try {
    const { start, end, rep, format } = req.query;
    const today = new Date().toISOString().slice(0,10);
    const params = [start || today, end || today];
    const repFilter = rep ? `AND rep_name = $3` : '';
    if (rep) params.push(rep);
    const { rows } = await pgPool.query(
      `SELECT
         rep_name,
         (logged_at AT TIME ZONE 'America/New_York')::date::text AS date,
         TO_CHAR(logged_at AT TIME ZONE 'America/New_York', 'HH12:MI AM') AS time,
         company_name, contact_name, contact_title,
         email, office_phone, mobile, card_address,
         location AS area_location,
         notable_moment AS notes,
         pm_opp, equip_opp, service_lead, piping_opp,
         sticker, sticker_count, vr_pres, appt_set, nothing
       FROM field_log_entries
       WHERE (logged_at AT TIME ZONE 'America/New_York')::date BETWEEN $1::date AND $2::date
       ${repFilter}
       ORDER BY date ASC, rep_name ASC, logged_at ASC`,
      params
    );

    if (format === 'xlsx') {
      function fmtDate(d) {
        if (!d) return '';
        const [y, m, day] = d.split('-');
        return `${m}/${day}/${y}`;
      }
      function teamOutcome(row) {
        const f = [];
        if (row.pm_opp)       f.push('Membership Opp');
        if (row.equip_opp)    f.push('Equipment Opp');
        if (row.service_lead) f.push('Service Lead');
        if (row.piping_opp)   f.push('Piping Opp');
        if (row.appt_set)     f.push('Appt Set');
        if (row.vr_pres)      f.push('VR Presentation');
        if (row.sticker)      f.push('Sticker Only');
        if (row.nothing)      f.push('Nothing Today');
        return f.join(', ');
      }
      const dataRows = rows.map(row => ({
        'Rep Name':      row.rep_name      || '',
        'Date':          fmtDate(row.date),
        'Time':          row.time          || '',
        'Company':       row.company_name  || '',
        'Contact':       row.contact_name  || '',
        'Title':         row.contact_title || '',
        'Email':         row.email         || '',
        'Mobile':        row.mobile        || '',
        'Phone':         row.office_phone  || '',
        'Address':       row.card_address  || '',
        'Branch':        '',
        'Outcome':       teamOutcome(row),
        'Sticker Count': row.sticker ? (row.sticker_count || 1) : 0,
        'Area/Location': row.area_location || '',
        'Notes':         row.notes         || '',
      }));
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(dataRows);
      XLSX.utils.book_append_sheet(wb, ws, 'Stops');
      const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="team_stops_export.xlsx"');
      return res.send(buf);
    }

    res.json(rows);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ─── REP LOCATION AUTOCOMPLETE ───────────────────────────────────────────────
app.get('/api/field-log/locations', async (req, res) => {
  if (!pgPool) return res.json([]);
  try {
    const { rep_name } = req.query;
    if (!rep_name) return res.json([]);
    const { rows } = await pgPool.query(
      `SELECT location, MAX(logged_at) AS last_used
       FROM field_log_entries
       WHERE rep_name = $1 AND location IS NOT NULL AND location != ''
       GROUP BY location
       ORDER BY last_used DESC
       LIMIT 50`,
      [rep_name]
    );
    res.json(rows.map(r => r.location));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ─── STOP SEARCH ─────────────────────────────────────────────────────────────
app.get('/api/field-log/search', async (req, res) => {
  if (!pgPool) return res.json([]);
  try {
    const { q } = req.query;
    if (!q || q.trim().length < 2) return res.json([]);
    const param = '%' + q.trim() + '%';
    const { rows } = await pgPool.query(
      `SELECT id, rep_name, company_name, contact_name, contact_title,
              email, mobile, office_phone, card_address, location,
              logged_at, sticker, sticker_count, notable_moment,
              pm_opp, equip_opp, service_lead, piping_opp, vr_pres, appt_set, nothing
       FROM field_log_entries
       WHERE company_name   ILIKE $1
          OR contact_name   ILIKE $1
          OR location       ILIKE $1
          OR notable_moment ILIKE $1
       ORDER BY logged_at DESC
       LIMIT 20`,
      [param]
    );
    function searchOutcome(row) {
      const f = [];
      if (row.pm_opp)       f.push('Membership Opp');
      if (row.equip_opp)    f.push('Equipment Opp');
      if (row.service_lead) f.push('Service Lead');
      if (row.piping_opp)   f.push('Piping Opp');
      if (row.appt_set)     f.push('Appt Set');
      if (row.vr_pres)      f.push('VR Presentation');
      if (row.sticker)      f.push('Sticker Only');
      if (row.nothing)      f.push('Nothing Today');
      return f.join(', ');
    }
    res.json(rows.map(r => ({
      id:           r.id,
      rep_name:     r.rep_name,
      company:      r.company_name    || '',
      contact_name: r.contact_name    || '',
      title:        r.contact_title   || '',
      email:        r.email           || '',
      mobile:       r.mobile          || '',
      phone:        r.office_phone    || '',
      address:      r.card_address    || '',
      area:         r.location        || '',
      logged_at:    r.logged_at,
      outcome:      searchOutcome(r),
      sticker_count: r.sticker ? (r.sticker_count || 1) : 0,
      notes:        r.notable_moment  || '',
    })));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ─── REP CSV EXPORT ───────────────────────────────────────────────────────────
app.get('/api/field-log/rep-export', async (req, res) => {
  if (!pgPool) return res.status(503).send('Database not configured');
  try {
    const { rep_name, range, from: fromDate, to: toDate, branch } = req.query;
    if (!rep_name) return res.status(400).send('rep_name required');
    const TZ = 'America/New_York';

    let dateFilter = '';
    const params = [rep_name];
    if (range === 'today') {
      dateFilter = `AND (logged_at AT TIME ZONE '${TZ}')::date = (NOW() AT TIME ZONE '${TZ}')::date`;
    } else if (range === 'week') {
      dateFilter = `AND logged_at >= DATE_TRUNC('week', NOW() AT TIME ZONE '${TZ}') AT TIME ZONE '${TZ}'`;
    } else if (range === 'month') {
      dateFilter = `AND logged_at >= DATE_TRUNC('month', NOW() AT TIME ZONE '${TZ}') AT TIME ZONE '${TZ}'`;
    } else if (range === 'custom' && fromDate && toDate) {
      dateFilter = `AND (logged_at AT TIME ZONE '${TZ}')::date BETWEEN $2::date AND $3::date`;
      params.push(fromDate, toDate);
    }

    const { rows } = await pgPool.query(
      `SELECT rep_name,
         TO_CHAR(logged_at AT TIME ZONE '${TZ}', 'MM/DD/YYYY') AS date,
         TO_CHAR(logged_at AT TIME ZONE '${TZ}', 'HH12:MI AM') AS time,
         company_name, contact_name, contact_title,
         email, mobile, office_phone, card_address,
         location AS area_location, notable_moment AS notes,
         pm_opp, equip_opp, service_lead, piping_opp,
         sticker, sticker_count, vr_pres, appt_set, nothing
       FROM field_log_entries
       WHERE rep_name = $1 ${dateFilter}
       ORDER BY logged_at ASC`,
      params
    );

    const esc = v => `"${String(v || '').replace(/"/g, '""')}"`;
    function outcome(row) {
      const f = [];
      if (row.pm_opp)       f.push('Membership Opp');
      if (row.equip_opp)    f.push('Equipment Opp');
      if (row.service_lead) f.push('Service Lead');
      if (row.piping_opp)   f.push('Piping Opp');
      if (row.appt_set)     f.push('Appt Set');
      if (row.vr_pres)      f.push('VR Presentation');
      if (row.sticker)      f.push('Sticker Only');
      if (row.nothing)      f.push('Nothing Today');
      return f.join(', ');
    }

    const { format } = req.query;
    const useXlsx = format !== 'csv';
    const today = new Date().toLocaleDateString('en-CA', { timeZone: TZ });
    const safeRep = rep_name.replace(/[^a-zA-Z0-9]/g, '_');
    const rangeLabel = range || 'all';

    const dataRows = rows.map(row => ({
      'Rep Name':      row.rep_name    || '',
      'Date':          row.date        || '',
      'Time':          row.time        || '',
      'Company':       row.company_name  || '',
      'Contact':       row.contact_name  || '',
      'Title':         row.contact_title || '',
      'Email':         row.email       || '',
      'Mobile':        row.mobile      || '',
      'Phone':         row.office_phone || '',
      'Address':       row.card_address || '',
      'Branch':        branch          || '',
      'Outcome':       outcome(row),
      'Sticker Count': row.sticker ? (row.sticker_count || 1) : 0,
      'Area/Location': row.area_location || '',
      'Notes':         row.notes       || '',
    }));

    if (useXlsx) {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(dataRows);
      XLSX.utils.book_append_sheet(wb, ws, 'Stops');
      const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
      const filename = `MyStops_${safeRep}_${rangeLabel}_${today}.xlsx`;
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
      res.send(buf);
    } else {
      const cols = ['Rep Name','Date','Time','Company','Contact','Title','Email','Mobile','Phone','Address','Branch','Outcome','Sticker Count','Area/Location','Notes'];
      const esc2 = v => `"${String(v ?? '').replace(/"/g, '""')}"`;
      const csvRows = dataRows.map(r => cols.map(c => typeof r[c] === 'number' ? r[c] : esc2(r[c])).join(','));
      const filename = `MyStops_${safeRep}_${rangeLabel}_${today}.csv`;
      res.setHeader('Content-Type', 'text/csv; charset=utf-8');
      res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
      res.send([cols.join(','), ...csvRows].join('\n'));
    }
  } catch (e) { res.status(500).send(e.message); }
});

// ─── REP GOALS ────────────────────────────────────────────────────────────────
app.get('/api/field-log/goals', async (req, res) => {
  if (!pgPool) return res.json([]);
  try {
    const { rows } = await pgPool.query('SELECT * FROM rep_goals ORDER BY rep_name');
    res.json(rows);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/field-log/goals', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  try {
    const { rep_name, wednesday_goal, daily_goal, weekly_goal } = req.body;
    if (!rep_name) return res.status(400).json({ error: 'rep_name required' });
    const { rows } = await pgPool.query(
      `INSERT INTO rep_goals (rep_name, wednesday_goal, daily_goal, weekly_goal, updated_at)
       VALUES ($1, $2, $3, $4, NOW())
       ON CONFLICT (rep_name) DO UPDATE SET
         wednesday_goal = EXCLUDED.wednesday_goal,
         daily_goal     = EXCLUDED.daily_goal,
         weekly_goal    = EXCLUDED.weekly_goal,
         updated_at     = NOW()
       RETURNING *`,
      [rep_name, wednesday_goal ?? 25, daily_goal ?? null, weekly_goal ?? 30]
    );
    res.json(rows[0]);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ─── SALES BLITZ ──────────────────────────────────────────────────────────────
// blitz_stops is SEPARATE from field_log_entries. Blitz logging pushes to
// monday.com only via the per-group rep assignment stored on the session; it
// never writes to field_log_entries.

// Map a stored blitz outcome string back to the boolean flags mondayUpsertProspect expects
function blitzOutcomeFlags(outcome) {
  const o = (outcome || '').toLowerCase();
  return {
    pmOpp:       o.includes('membership'),
    equipOpp:    o.includes('equipment'),
    serviceLead: o.includes('service lead'),
    pipingOpp:   o.includes('piping'),
    sticker:     o.includes('sticker'),
    vrPres:      o.includes('vr pres'),
    apptSet:     o.includes('appointment'),
    nothing:     o.trim() === 'nothing' || o.includes('nothing today'),
  };
}

// The monday.com Branch dropdown options are bare city names — return just the
// city so it matches a dropdown label exactly (passed as { labels: [branch] }).
const BLITZ_BRANCH_CITIES = ['Atlanta','Charlotte','Greenville','Tampa','Nashville','Cleveland','Detroit','Chicago','Dallas'];
function territoryToBranch(territory) {
  if (!territory) return null;
  const t = territory.toLowerCase();
  const city = BLITZ_BRANCH_CITIES.find(c => t.includes(c.toLowerCase()));
  return city ? city : null;
}

app.post('/api/blitz/upload', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  const { name, stops, group_assignments } = req.body;
  if (!name || !Array.isArray(stops) || stops.length === 0) {
    return res.status(400).json({ error: 'name and stops[] required' });
  }
  const num = v => { const n = Number(v); return (v == null || v === '' || isNaN(n)) ? null : n; };
  const client = await pgPool.connect();
  try {
    await client.query('BEGIN');
    // Archive (do NOT delete) any currently active session — its stops stay queryable
    await client.query(`UPDATE blitz_sessions SET active=FALSE WHERE active=TRUE`);
    const s = await client.query(
      `INSERT INTO blitz_sessions (name, active, group_assignments) VALUES ($1, TRUE, $2) RETURNING id`,
      [name, JSON.stringify(group_assignments || {})]
    );
    const sessionId = s.rows[0].id;
    let inserted = 0;
    for (const st of stops) {
      const company = (st.company || '').toString().trim();
      const group   = (st.group_name || '').toString().trim();
      if (!company || !group) continue;
      const day  = parseInt(String(st.day).replace(/[^0-9]/g, ''), 10) || 1;
      const role = (st.role || 'PRIMARY').toString().trim() || 'PRIMARY';
      const stopNumber = (st.stop_number != null ? String(st.stop_number) : '').trim() || '—';
      await client.query(
        `INSERT INTO blitz_stops
           (session_id, group_name, day, stop_number, role, company, address, miles_from_shop,
            industry, contact_name, title, phone, air_likelihood, site_confidence, territory,
            priority_score, website, contact_email)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18)`,
        [sessionId, group, day, stopNumber, role, company,
         st.address || null, num(st.miles_from_shop), st.industry || null,
         st.contact_name || null, st.title || null, st.phone || null,
         st.air_likelihood || null, st.site_confidence || null, st.territory || null,
         num(st.priority_score), st.website || null, st.contact_email || null]
      );
      inserted++;
    }
    await client.query('COMMIT');
    res.json({ session_id: sessionId, totalStops: inserted });
  } catch (e) {
    await client.query('ROLLBACK');
    console.error('blitz upload error:', e.message);
    res.status(500).json({ error: e.message });
  } finally {
    client.release();
  }
});

app.get('/api/blitz/active', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  try {
    const s = await pgPool.query(
      `SELECT id, name, created_at, active, group_assignments FROM blitz_sessions WHERE active=TRUE ORDER BY id DESC LIMIT 1`
    );
    if (s.rows.length === 0) return res.json({ session: null, stops: [] });
    const session = s.rows[0];
    const st = await pgPool.query(
      `SELECT * FROM blitz_stops WHERE session_id=$1
       ORDER BY group_name ASC, day ASC,
                NULLIF(regexp_replace(stop_number, '[^0-9]', '', 'g'), '')::int NULLS LAST,
                id ASC`,
      [session.id]
    );
    res.json({ session, stops: st.rows });
  } catch (e) {
    console.error('blitz active error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/blitz/log/:stopId', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  const id = parseInt(req.params.stopId, 10);
  if (!id) return res.status(400).json({ error: 'Invalid stop id' });
  const { outcome, notes, repName } = req.body;
  try {
    const { rows } = await pgPool.query(
      `UPDATE blitz_stops SET logged=TRUE, logged_at=NOW(), outcome=$1, notes=$2
       WHERE id=$3 RETURNING *`,
      [outcome || null, notes || null, id]
    );
    if (rows.length === 0) return res.status(404).json({ error: 'Stop not found' });
    const stop = rows[0];
    res.json(stop);

    // Fire-and-forget monday.com push using the group's assigned rep,
    // falling back to the logging rep if the group has no assignment
    if (process.env.MONDAY_API_KEY && stop.company) {
      (async () => {
        try {
          const sres = await pgPool.query(
            `SELECT group_assignments FROM blitz_sessions WHERE id=$1`, [stop.session_id]
          );
          const ga = (sres.rows[0] && sres.rows[0].group_assignments) || {};
          const assignedRep = ga[stop.group_name] || repName || null;
          if (!assignedRep) return;
          // Morty and Hudson are not tracked in monday.com — never push their stops
          if (assignedRep === 'Morty' || assignedRep === 'Hudson') return;
          mondayUpsertProspect({
            repName:       assignedRep,
            companyName:   stop.company,
            contactName:   stop.contact_name,
            location:      stop.territory,
            notableMoment: stop.notes,
            ...blitzOutcomeFlags(stop.outcome),
            stickerCount:  0,
            mobile:        stop.mobile || stop.phone,
            officePhone:   stop.phone,
            email:         stop.email || stop.contact_email,
            branch:        territoryToBranch(stop.territory),
          }).catch(e => console.error('blitz monday upsert error:', e.message));
        } catch (e) { console.error('blitz monday push error:', e.message); }
      })();
    }
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/blitz/log/:stopId', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  const id = parseInt(req.params.stopId, 10);
  if (!id) return res.status(400).json({ error: 'Invalid stop id' });
  try {
    const { rows } = await pgPool.query(
      `UPDATE blitz_stops SET logged=FALSE, logged_at=NULL, outcome=NULL, notes=NULL
       WHERE id=$1 RETURNING id`, [id]
    );
    if (rows.length === 0) return res.status(404).json({ error: 'Stop not found' });
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/blitz/stops/manual', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  const { session_id, group_name, day, company, contact_name, title, phone, mobile, email, website, address, notes } = req.body;
  if (!session_id || !group_name || !company) {
    return res.status(400).json({ error: 'session_id, group_name and company required' });
  }
  try {
    const mx = await pgPool.query(
      `SELECT stop_number FROM blitz_stops WHERE session_id=$1 AND role='MANUAL'`, [session_id]
    );
    let maxN = 0;
    mx.rows.forEach(r => { const m = String(r.stop_number).match(/^M(\d+)$/i); if (m) maxN = Math.max(maxN, parseInt(m[1], 10)); });
    const stopNumber = 'M' + (maxN + 1);
    const d = parseInt(String(day).replace(/[^0-9]/g, ''), 10) || 1;
    const { rows } = await pgPool.query(
      `INSERT INTO blitz_stops
         (session_id, group_name, day, stop_number, role, company, contact_name, title, phone, mobile, email, website, address, notes)
       VALUES ($1,$2,$3,$4,'MANUAL',$5,$6,$7,$8,$9,$10,$11,$12,$13) RETURNING *`,
      [session_id, group_name, d, stopNumber, company.trim(),
       contact_name || null, title || null, phone || null, mobile || null,
       email || null, website || null, address || null, notes || null]
    );
    res.json(rows[0]);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.patch('/api/blitz/stops/:stopId/contact', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  const id = parseInt(req.params.stopId, 10);
  if (!id) return res.status(400).json({ error: 'Invalid stop id' });
  const { contact_name, title, phone, email, mobile, website, address } = req.body;
  try {
    const { rows } = await pgPool.query(
      `UPDATE blitz_stops SET
         contact_name = $1, title = $2, phone = $3, email = $4,
         mobile = $5, website = $6, address = $7
       WHERE id = $8 RETURNING *`,
      [contact_name || null, title || null, phone || null, email || null,
       mobile || null, website || null, address || null, id]
    );
    if (rows.length === 0) return res.status(404).json({ error: 'Stop not found' });
    res.json(rows[0]);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/blitz/sessions', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  try {
    const { rows } = await pgPool.query(
      `SELECT s.id, s.name, s.created_at, s.active,
              COUNT(bs.id)::int AS "totalStops",
              COUNT(bs.id) FILTER (WHERE bs.logged)::int AS "loggedStops"
       FROM blitz_sessions s
       LEFT JOIN blitz_stops bs ON bs.session_id = s.id
       GROUP BY s.id
       ORDER BY s.id DESC`
    );
    res.json(rows);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/blitz/export/:sessionId', async (req, res) => {
  if (!pgPool) return res.status(503).send('Database not configured');
  const sid = parseInt(req.params.sessionId, 10);
  if (!sid) return res.status(400).send('Invalid session id');
  try {
    const sres = await pgPool.query(`SELECT name, group_assignments FROM blitz_sessions WHERE id=$1`, [sid]);
    if (sres.rows.length === 0) return res.status(404).send('Session not found');
    const ga = sres.rows[0].group_assignments || {};
    const { rows } = await pgPool.query(
      `SELECT * FROM blitz_stops WHERE session_id=$1
       ORDER BY group_name, day,
                NULLIF(regexp_replace(stop_number, '[^0-9]', '', 'g'), '')::int NULLS LAST, id`,
      [sid]
    );
    const data = rows.map(r => ({
      'Group':           r.group_name,
      'Day':             r.day,
      'Stop':            r.stop_number,
      'Role':            r.role,
      'Company':         r.company,
      'Address':         r.address || '',
      'Territory':       r.territory || '',
      'Industry':        r.industry || '',
      'Contact':         r.contact_name || '',
      'Title':           r.title || '',
      'Phone':           r.phone || '',
      'Air Likelihood':  r.air_likelihood || '',
      'Site Confidence': r.site_confidence || '',
      'Priority Score':  r.priority_score != null ? Number(r.priority_score) : '',
      'Logged':          r.logged ? 'Yes' : 'No',
      'Logged At':       r.logged_at ? new Date(r.logged_at).toLocaleString('en-US', { timeZone: 'America/New_York' }) : '',
      'Outcome':         r.outcome === 'nothing' ? 'Nothing Today' : (r.outcome || ''),
      'Notes':           r.notes || '',
      'Assigned Rep':    ga[r.group_name] || '',
    }));
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, 'Blitz Stops');
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const safe = String(sres.rows[0].name || 'blitz').replace(/[^a-zA-Z0-9]/g, '_').slice(0, 40);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="Blitz_${safe}_${sid}.xlsx"`);
    res.send(buf);
  } catch (e) { res.status(500).send(e.message); }
});

app.get('/api/blitz/dashboard', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  try {
    const s = await pgPool.query(
      `SELECT id FROM blitz_sessions WHERE active=TRUE ORDER BY id DESC LIMIT 1`
    );
    if (s.rows.length === 0) return res.json({ session_id: null, groups: [], totals: { stops: 0, opps: 0 } });
    const sessionId = s.rows[0].id;
    const { rows } = await pgPool.query(
      `SELECT group_name AS "groupName", day,
              COUNT(*) FILTER (WHERE logged)::int AS stops,
              COUNT(*) FILTER (WHERE logged AND outcome IS NOT NULL AND lower(btrim(outcome)) <> 'nothing')::int AS opps
       FROM blitz_stops WHERE session_id=$1
       GROUP BY group_name, day
       ORDER BY group_name, day`,
      [sessionId]
    );
    const totals = rows.reduce((a, r) => ({ stops: a.stops + r.stops, opps: a.opps + r.opps }), { stops: 0, opps: 0 });
    res.json({ session_id: sessionId, groups: rows, totals });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/blitz/session', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  try {
    await pgPool.query(`UPDATE blitz_sessions SET active=FALSE WHERE active=TRUE`);
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ─── SLACK FEEDBACK ───────────────────────────────────────────────────────────
app.post('/api/slack-feedback', async (req, res) => {
  const { rep, branch, type, message, tab } = req.body;
  const webhookUrl = process.env.SLACK_WEBHOOK_URL || 'SLACK_WEBHOOK_PLACEHOLDER';
  if (webhookUrl === 'SLACK_WEBHOOK_PLACEHOLDER') {
    return res.status(503).json({ error: 'Slack webhook not configured' });
  }
  const text = [
    '🛠️ *AirIQ Feedback*',
    `*From:* ${rep || 'Unknown'} — ${branch || 'Unknown'}`,
    `*Type:* ${type || 'General'}`,
    `*Message:* ${message || '(no message)'}`,
    `*Tab:* ${tab || 'Unknown'}`,
    `*Time:* ${new Date().toLocaleString('en-US', { timeZone: 'America/New_York' })} EST`,
  ].join('\n');
  try {
    const r = await fetch(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text }),
    });
    if (!r.ok) throw new Error(`Slack ${r.status}`);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/', (req, res) => {
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
  res.sendFile(path.join(__dirname, 'index.html'));
});

// ─── STARTUP & SCHEDULER ─────────────────────────────────────────────────────

const PORT = process.env.PORT || 3000;
app.listen(PORT, async () => {
  console.log(`AirIQ running on port ${PORT}`);
  try { await initFieldLogTable(); } catch (e) { console.error('Field log table init failed:', e.message); }
  try { await initBlitzTables(); } catch (e) { console.error('Blitz table init failed:', e.message); }
  try {
    await fetchInventoryFromGmail();
    if (inventoryCache.length === 0) loadInventoryFallback();
  } catch (e) {
    console.error('Initial Gmail inventory load failed:', e.message);
    loadInventoryFallback();
  }
  if (process.env.MONDAY_API_KEY) {
    try { await fetchMondayMeta(); } catch (e) { console.error('Monday.com meta fetch failed:', e.message); }
    try { await fetchMondayUsers(); } catch (e) { console.error('Monday.com users fetch failed:', e.message); }
  }
  // Start health check — run once immediately then every 5 minutes
  runHealthCheck().catch(e => console.error('Initial health check failed:', e.message));
  setInterval(() => runHealthCheck().catch(e => console.error('Health check error:', e.message)), 5 * 60 * 1000);
});

// Refresh daily at 6:10am EST (11:10 UTC) — after the 5:56am ST email arrives
function scheduleDailyRefresh() {
  const now = new Date();
  const next = new Date();
  next.setUTCHours(11, 10, 0, 0);
  if (next <= now) next.setUTCDate(next.getUTCDate() + 1);
  const msUntilNext = next - now;
  console.log(`Next inventory refresh scheduled in ${Math.round(msUntilNext / 60000)} minutes`);
  setTimeout(async () => {
    try {
      await fetchInventoryFromGmail();
    } catch (e) {
      console.error('Scheduled Gmail inventory refresh failed:', e.message);
    }
    scheduleDailyRefresh(); // reschedule for the next day
  }, msUntilNext);
}

scheduleDailyRefresh();
