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
  await pgPool.query(`ALTER TABLE field_log_entries ADD COLUMN IF NOT EXISTS card_address  TEXT`);
  await pgPool.query(`CREATE INDEX IF NOT EXISTS idx_field_log_rep_logged ON field_log_entries (rep_name, logged_at DESC)`);
  console.log('field_log_entries table ready');
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

// In-memory inventory cache
let inventoryCache = [];
let lastCacheUpdate = null;
let isFetching = false;

// ─── GMAIL INVENTORY LOADER ───────────────────────────────────────────────────
// Reads the daily "Hodge Compressor Inventory" email from Gmail,
// downloads the .xlsx attachment, and populates the inventory cache.
// No ST API calls needed — zero rate limit risk.

async function getGmailToken() {
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id:     process.env.GMAIL_CLIENT_ID,
      client_secret: process.env.GMAIL_CLIENT_SECRET,
      refresh_token: process.env.GMAIL_REFRESH_TOKEN,
      grant_type:    'refresh_token',
    }),
  });
  const data = await res.json();
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
        headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
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
    const token = await getGmailToken();

    // Step 1: Search for the most recent Hodge Compressor Inventory email
    const searchRes = await fetch(
      'https://gmail.googleapis.com/gmail/v1/users/me/messages?maxResults=1&q=subject:"Hodge+Compressor+Inventory"+from:noreply@onservicetitan.com',
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const searchData = await searchRes.json();

    if (!searchData.messages || searchData.messages.length === 0) {
      console.error('No Hodge Compressor Inventory email found in Gmail');
      return;
    }

    const messageId = searchData.messages[0].id;
    console.log(`Found inventory email: ${messageId}`);

    // Step 2: Get the message to find the attachment ID
    const msgRes = await fetch(
      `https://gmail.googleapis.com/gmail/v1/users/me/messages/${messageId}`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const msgData = await msgRes.json();

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

    console.log(`Downloading attachment: ${filename}`);

    // Step 3: Download the attachment
    const attachRes = await fetch(
      `https://gmail.googleapis.com/gmail/v1/users/me/messages/${messageId}/attachments/${attachmentId}`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const attachData = await attachRes.json();

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
    console.log(`Inventory cache loaded from Gmail: ${parsed.length} equipment items at ${lastCacheUpdate}`);

    if (parsed.length > 0) {
      console.log('Sample item:', JSON.stringify(parsed[0]));
    }

  } catch (err) {
    console.error('fetchInventoryFromGmail error:', err.message);
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
  try {
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': process.env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
        'anthropic-beta': 'tools-2024-04-04'
      },
      body: JSON.stringify(req.body)
    });
    const data = await response.json();
    res.json(data);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ST pricebook proxy (unchanged)
app.post('/api/st', async (req, res) => {
  try {
    const response = await fetch(PROXY_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(req.body)
    });
    const data = await response.json();
    res.json(data);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Inventory lookup from cache
app.get('/api/inventory', (req, res) => {
  const term = (req.query.q || '').trim();
  if (!term) {
    return res.json({ items: [], cacheAge: lastCacheUpdate, totalCached: inventoryCache.length });
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

  res.json({ items: results, cacheAge: lastCacheUpdate, totalCached: inventoryCache.length });
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

app.get('/api/sheets/setup', async (req, res) => {
  try {
    const token = await getGmailToken();
    const range = encodeURIComponent(`${SHEET_NAME}!A1`);
    const r = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${range}?valueInputOption=USER_ENTERED`,
      {
        method: 'PUT',
        headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
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

// ─── HEALTH CHECK ────────────────────────────────────────────────────────────
app.get('/api/health', async (req, res) => {
  if (!pgPool) return res.json({ status: 'ok', db: 'not configured', timestamp: new Date() });
  try {
    await pgPool.query('SELECT 1');
    res.json({ status: 'ok', db: 'connected', timestamp: new Date() });
  } catch (e) {
    console.error('Health check DB error:', e.message);
    res.status(503).json({ status: 'error', db: 'disconnected', error: e.message });
  }
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
  if (faqCache.items && (Date.now() - faqCache.fetchedAt) < 10 * 60 * 1000) {
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
  'Jon Sanders': 'Jon Sanders',
};
let mondayColMap  = null;
let mondayGroupId = null;
let mondayUserMap = {};

async function mondayGraphQL(query, variables) {
  const body = variables ? { query, variables } : { query };
  const res = await fetch('https://api.monday.com/v2', {
    method: 'POST',
    headers: {
      'Content-Type':  'application/json',
      'Authorization': `Bearer ${process.env.MONDAY_API_KEY}`,
    },
    body: JSON.stringify(body),
  });
  const data = await res.json();
  if (data.errors) throw new Error(data.errors.map(e => e.message).join('; '));
  return data;
}

async function fetchMondayMeta() {
  const data = await mondayGraphQL(`
    query {
      boards(ids: [${MONDAY_BOARD_ID}]) {
        columns { id title }
        groups  { id title }
      }
    }
  `);
  const board = data.data?.boards?.[0];
  if (!board) throw new Error('Monday.com board not found');

  const byTitle = {};
  board.columns.forEach(c => { byTitle[c.title] = c.id; });

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
            apptSet, nothing, stickerCount, mobile, officePhone, email } = stopData;

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
    const searchRaw = await fetch('https://api.monday.com/v2', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${process.env.MONDAY_API_KEY}`,
        'Content-Type': 'application/json',
        'API-Version': '2024-01',
      },
      body: JSON.stringify({ query: searchQuery }),
    });
    const searchData = await searchRaw.json();
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
      if (mondayColMap.phone       && phoneVal)      colValues[mondayColMap.phone]       = phoneVal;
      if (mondayColMap.email       && email)         colValues[mondayColMap.email]       = email;
      if (mondayColMap.location    && location)      colValues[mondayColMap.location]    = location;
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
      if (mondayColMap.rep && mondayUserId)           colValues[mondayColMap.rep]         = { personsAndTeams: [{ id: mondayUserId, kind: 'person' }] };
      if (mondayColMap.status)                       colValues[mondayColMap.status]      = { label: 'Active Prospects' };
      if (mondayColMap.lastVisited)                  colValues[mondayColMap.lastVisited] = { date: today };
      if (mondayColMap.outcome     && outcomeStr)    colValues[mondayColMap.outcome]     = outcomeStr;
      if (mondayColMap.notes       && notableMoment) colValues[mondayColMap.notes]       = notableMoment;
      if (mondayColMap.contactName && contactName)   colValues[mondayColMap.contactName] = contactName;
      const phoneVal = mobile || officePhone;
      if (mondayColMap.phone       && phoneVal)      colValues[mondayColMap.phone]       = phoneVal;
      if (mondayColMap.email       && email)         colValues[mondayColMap.email]       = email;
      if (mondayColMap.location    && location)      colValues[mondayColMap.location]    = location;
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
            location, notableMoment, companyName, contactName, stickerCount,
            mobile, officePhone, email, website, cardAddress } = req.body;
    if (!repName) return res.status(400).json({ error: 'repName required' });
    const sc = sticker ? Math.max(1, Number(stickerCount) || 1) : 0;
    const { rows } = await pgPool.query(
      `INSERT INTO field_log_entries
         (rep_name, pm_opp, equip_opp, service_lead, piping_opp, sticker, vr_pres, appt_set, nothing,
          location, notable_moment, company_name, contact_name, sticker_count,
          mobile, office_phone, email, website, card_address)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19)
       RETURNING id, logged_at`,
      [repName, !!pmOpp, !!equipOpp, !!serviceLead, !!pipingOpp, !!sticker, !!vrPres, !!apptSet, !!nothing,
       location || null, notableMoment || null, companyName || null, contactName || null, sc,
       mobile || null, officePhone || null, email || null, website || null, cardAddress || null]
    );
    res.json({ success: true, id: rows[0].id, loggedAt: rows[0].logged_at });
    if (companyName && process.env.MONDAY_API_KEY) {
      mondayUpsertProspect({
        repName, companyName, contactName, location, notableMoment,
        pmOpp: !!pmOpp, equipOpp: !!equipOpp, serviceLead: !!serviceLead,
        pipingOpp: !!pipingOpp, sticker: !!sticker, vrPres: !!vrPres,
        apptSet: !!apptSet, nothing: !!nothing,
        stickerCount: sc, mobile, officePhone, email,
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
       WHERE rep_name = $1 AND logged_at >= NOW()::date`,
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
         MAX(logged_at)              AS last_logged
       FROM field_log_entries
       WHERE logged_at >= NOW()::date
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
       WHERE logged_at >= DATE_TRUNC('week', NOW() AT TIME ZONE 'America/New_York')
       GROUP BY 1
       ORDER BY 1`
    );
    res.json(rows);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/field-log/stops/:repName', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  try {
    const { rows } = await pgPool.query(
      `SELECT id, logged_at, pm_opp, equip_opp, service_lead, piping_opp, sticker, vr_pres, appt_set, nothing,
              location, notable_moment, company_name, contact_name, sticker_count,
              mobile, office_phone, email, website, card_address
       FROM field_log_entries
       WHERE rep_name = $1 AND logged_at >= NOW()::date
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
    let dateFilter = '';
    let limitClause = 'LIMIT 2000';
    if (range === 'today') {
      dateFilter = `AND logged_at >= NOW()::date`;
      limitClause = '';
    } else if (range === 'week') {
      dateFilter = `AND logged_at >= DATE_TRUNC('week', NOW() AT TIME ZONE 'America/New_York')`;
      limitClause = '';
    } else if (range === 'month') {
      dateFilter = `AND logged_at >= DATE_TRUNC('month', NOW() AT TIME ZONE 'America/New_York')`;
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
    const { start, end } = req.query;
    const today = new Date().toISOString().slice(0,10);
    const { rows } = await pgPool.query(
      `SELECT
         rep_name,
         (logged_at AT TIME ZONE 'America/New_York')::date::text AS date,
         COUNT(*)::int                                                        AS total_stops,
         SUM(pm_opp::int)::int                                                AS pm_opps,
         SUM(equip_opp::int)::int                                             AS equip_opps,
         SUM(service_lead::int)::int                                          AS service_leads,
         SUM(piping_opp::int)::int                                            AS piping_opps,
         SUM(COALESCE(sticker_count, sticker::int))::int                        AS stickers,
         SUM(vr_pres::int)::int                                               AS vr_pres,
         SUM(appt_set::int)::int                                              AS appt_set,
         SUM(nothing::int)::int                                               AS nothing,
         STRING_AGG(DISTINCT location, '; ')     FILTER (WHERE location     IS NOT NULL AND location     <> '') AS locations,
         STRING_AGG(notable_moment, ' | ')        FILTER (WHERE notable_moment IS NOT NULL AND notable_moment <> '') AS notable_moments,
         STRING_AGG(DISTINCT company_name, '; ') FILTER (WHERE company_name IS NOT NULL AND company_name <> '') AS companies,
         STRING_AGG(DISTINCT contact_name, '; ') FILTER (WHERE contact_name IS NOT NULL AND contact_name <> '') AS contacts
       FROM (
         SELECT *, (logged_at AT TIME ZONE 'America/New_York')::date AS date_local
         FROM field_log_entries
         WHERE (logged_at AT TIME ZONE 'America/New_York')::date BETWEEN $1::date AND $2::date
       ) sub
       GROUP BY rep_name, date
       ORDER BY date, rep_name`,
      [start || today, end || today]
    );
    res.json(rows);
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
  try { await fetchInventoryFromGmail(); } catch (e) { console.error('Initial Gmail inventory load failed:', e.message); }
  if (process.env.MONDAY_API_KEY) {
    try { await fetchMondayMeta(); } catch (e) { console.error('Monday.com meta fetch failed:', e.message); }
    try { await fetchMondayUsers(); } catch (e) { console.error('Monday.com users fetch failed:', e.message); }
  }
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
