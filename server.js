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
  ? new Pool({ connectionString: process.env.DATABASE_URL, ssl: { rejectUnauthorized: false } })
  : null;

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
  'Drawing Checklist Complete','Manager Approval','Notes','Submitted At',
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

// ─── FIELD LOG ────────────────────────────────────────────────────────────────
app.post('/api/field-log', async (req, res) => {
  if (!pgPool) return res.status(503).json({ error: 'Database not configured' });
  try {
    const { repName, pmOpp, equipOpp, serviceLead, pipingOpp, sticker, vrPres, apptSet, nothing, location, notableMoment } = req.body;
    if (!repName) return res.status(400).json({ error: 'repName required' });
    const { rows } = await pgPool.query(
      `INSERT INTO field_log_entries
         (rep_name, pm_opp, equip_opp, service_lead, piping_opp, sticker, vr_pres, appt_set, nothing, location, notable_moment)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11)
       RETURNING id, logged_at`,
      [repName, !!pmOpp, !!equipOpp, !!serviceLead, !!pipingOpp, !!sticker, !!vrPres, !!apptSet, !!nothing,
       location || null, notableMoment || null]
    );
    res.json({ success: true, id: rows[0].id, loggedAt: rows[0].logged_at });
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
         SUM(sticker::int)::int      AS stickers,
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
         SUM(appt_set::int)::int     AS appt_set,
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
