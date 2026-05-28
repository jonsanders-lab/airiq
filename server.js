const express = require('express');
const fetch = require('node-fetch');
const path = require('path');
const XLSX = require('xlsx');
const app = require('express')();

app.use(express.json());
app.use(express.static('.'));

const PROXY_URL = 'https://airiq-st-proxy.jon-sanders.workers.dev';

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

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// ─── STARTUP & SCHEDULER ─────────────────────────────────────────────────────

const PORT = process.env.PORT || 3000;
app.listen(PORT, async () => {
  console.log(`AirIQ running on port ${PORT}`);
  // Load inventory from today's Gmail report on startup
  try {
    await fetchInventoryFromGmail();
  } catch (e) {
    console.error('Initial Gmail inventory load failed:', e.message);
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
