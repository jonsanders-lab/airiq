const express = require('express');
const fetch = require('node-fetch');
const path = require('path');
const app = express();

app.use(express.json());
app.use(express.static('.'));

const PROXY_URL = 'https://airiq-st-proxy.jon-sanders.workers.dev';

// ST credentials — set these in Railway's Variables tab, never hardcode
const ST_TENANT        = process.env.ST_TENANT;
const ST_APP_KEY       = process.env.ST_APP_KEY;
const ST_CLIENT_ID     = process.env.ST_CLIENT_ID;
const ST_CLIENT_SECRET = process.env.ST_CLIENT_SECRET;
const INVENTORY_REPORT_ID = 1823;

// Hodge Compressor warehouse — the only location AirIQ needs
// ID sourced from ST inventory-locations/warehouses/details/84363890
const INVENTORY_LOCATION_IDS = [
  84363890,  // Hodge Compressor
  11053988,  // Nashville Warehouse
  11055398,  // Charlotte Warehouse
  11055907,  // Atlanta Warehouse
  42353491,  // Tampa Warehouse
  42353988,  // Greenville Warehouse
  46238016,  // Cleveland Warehouse
  73998548,  // Detroit Inventory
  73998551,  // Nashville Inventory
  73998799,  // Atlanta Inventory
  73998802,  // Cleveland Inventory
  74001225,  // Charlotte Inventory
  74001353,  // Greenville Inventory
  74001481,  // Tampa Inventory
  78073441,  // Dallas Compressor
  78078793,  // Dallas Inventory
  32766265,  // Detroit Compressor
  108478665, // Chicago Warehouse
  108716118  // Chicago Inventory
];

// In-memory inventory cache
let inventoryCache = [];
let lastCacheUpdate = null;
let isFetching = false;

// Fetch ST access token
async function getSTToken() {
  const res = await fetch("https://auth.servicetitan.io/connect/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "client_credentials",
      client_id: ST_CLIENT_ID,
      client_secret: ST_CLIENT_SECRET,
    }),
  });
  const data = await res.json();
  return data.access_token;
}

// Fetch all pages of the Inventory Stock Report
async function fetchInventoryReport() {
  if (isFetching) {
    console.log("Fetch already in progress, skipping...");
    return;
  }
  isFetching = true;
  console.log("Fetching inventory report from ST...");

  try {
    let token = await getSTToken();
    let page = 1;
    let allRows = [];
    let hasMore = true;

    while (hasMore) {
      const res = await fetch(
        `https://api.servicetitan.io/reporting/v2/tenant/${ST_TENANT}/report-category/inventory/reports/${INVENTORY_REPORT_ID}/data`,
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${token}`,
            "ST-App-Key": ST_APP_KEY,
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            page,
            pageSize: 1000,
            parameters: [
              {
                name: "Date",
                value: new Date().toISOString().split('T')[0]  // today's date e.g. "2026-05-27"
              },
              {
                name: "InventoryLocationIds",
                value: INVENTORY_LOCATION_IDS
              }
            ]
          })
        }
      );

      if (!res.ok) {
        if (res.status === 429) {
          console.log("Rate limited, waiting 30 seconds...");
          await new Promise(resolve => setTimeout(resolve, 30000));
          continue;
        }
        if (res.status === 401) {
          console.log("Token expired, refreshing...");
          token = await getSTToken();
          continue;
        }
        const errBody = await res.text();
        console.error(`ST inventory report error ${res.status}:`, errBody);
        break;
      }

      const data = await res.json();
      const rows = data.data || [];
      allRows = allRows.concat(rows);
      hasMore = data.hasMore;
      page++;
      console.log(`Fetched page ${page - 1}, got ${rows.length} rows, hasMore: ${hasMore}`);

      await new Promise(resolve => setTimeout(resolve, 5000));
    }

    if (allRows.length > 0) {
      console.log("Sample row:", JSON.stringify(allRows[0]));
    }

    const EQUIPMENT_PREFIXES = ['HD','HB','HV','HT','HAD','HTM'];

    const parsed = allRows
      .filter(row => {
        const code = (row[1] || '').toUpperCase();
        return EQUIPMENT_PREFIXES.some(p => code.startsWith(p));
      })
      .map(row => ({
        name: row[0],              // ItemName
        code: row[1],              // ItemCode
        description: row[2],       // ItemDescription
        type: row[3],              // ItemType
        location: row[4],          // InventoryLocation
        quantityAvailable: row[5], // QuantityAvailable
        quantityOnHold: row[6],    // QuantityOnHold
        quantityOnOrder: row[7],   // QuantityOnOrder
        quantityOnHand: row[8],    // QuantityOnHand
        binLocation: row[11],      // BinLocation
        quantityReserved: row[12], // QuantityReserved
      }));

    inventoryCache = parsed;
    lastCacheUpdate = new Date().toISOString();
    console.log(`Inventory cache updated: ${parsed.length} equipment items at ${lastCacheUpdate}`);
  } catch (err) {
    console.error("fetchInventoryReport error:", err.message);
  } finally {
    isFetching = false;
  }
}

// Search inventory cache by item code or name
function searchInventoryCache(searchTerm) {
  const term = searchTerm.toUpperCase();
  return inventoryCache.filter(item =>
    (item.code && item.code.toUpperCase().startsWith(term)) ||
    (item.name && item.name.toUpperCase().startsWith(term))
  );
}

// Aggregate inventory across warehouse locations
function aggregateInventory(rows) {
  const warehouseLocations = ['inventory', 'warehouse'];
  const warehouseRows = rows.filter(r =>
    r.location && warehouseLocations.some(w => r.location.toLowerCase().includes(w))
  );

  if (warehouseRows.length === 0) return rows;

  const totals = {
    quantityOnHand: 0,
    quantityAvailable: 0,
    quantityOnOrder: 0,
    quantityReserved: 0,
    locations: []
  };

  warehouseRows.forEach(row => {
    totals.quantityOnHand += (row.quantityOnHand || 0);
    totals.quantityAvailable += (row.quantityAvailable || 0);
    totals.quantityOnOrder += (row.quantityOnOrder || 0);
    totals.quantityReserved += (row.quantityReserved || 0);
    if ((row.quantityOnHand || 0) > 0 || (row.quantityAvailable || 0) > 0) {
      totals.locations.push({
        location: row.location,
        onHand: row.quantityOnHand,
        available: row.quantityAvailable,
        onOrder: row.quantityOnOrder,
      });
    }
  });

  return totals;
}

// API: Claude proxy
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

// API: ST proxy (pricebook)
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

// DIAGNOSTIC: Fetch report schema to discover required parameters
app.get('/api/inventory/report-schema', async (req, res) => {
  try {
    const token = await getSTToken();
    const response = await fetch(
      `https://api.servicetitan.io/reporting/v2/tenant/${ST_TENANT}/report-category/inventory/reports/${INVENTORY_REPORT_ID}`,
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`,
          "ST-App-Key": ST_APP_KEY,
        }
      }
    );
    const status = response.status;
    const body = await response.json();
    res.json({ httpStatus: status, reportSchema: body });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// API: Inventory lookup from cache
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
    code: item.code,
    name: item.name,
    inventory: aggregateInventory(item.rows),
    rawRows: item.rows,
  }));

  res.json({ items: results, cacheAge: lastCacheUpdate, totalCached: inventoryCache.length });
});

// API: Force inventory refresh
app.post('/api/inventory/refresh', async (req, res) => {
  try {
    await fetchInventoryReport();
    res.json({ success: true, totalRows: inventoryCache.length, updatedAt: lastCacheUpdate });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// API: Cache status
app.get('/api/inventory/status', (req, res) => {
  res.json({ totalRows: inventoryCache.length, lastUpdated: lastCacheUpdate, isFetching });
});

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, async () => {
  console.log(`AirIQ running on port ${PORT}`);
  try {
    await fetchInventoryReport();
  } catch (e) {
    console.error("Initial inventory fetch failed:", e.message);
  }
});

// Refresh inventory every 4 hours
setInterval(async () => {
  try {
    await fetchInventoryReport();
  } catch (e) {
    console.error("Inventory refresh failed:", e.message);
  }
}, 4 * 60 * 60 * 1000);
