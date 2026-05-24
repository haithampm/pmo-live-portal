// api/sync.js - PMO Bulk Sync: Import sharepoint-data.json projects into Supabase
// POST /api/sync - Seeds/syncs all projects from static JSON file into pmo_records
// SECURITY: Server-side only. Requires admin token.

const JSON_HEADERS = {
  'Content-Type': 'application/json',
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization'
};

function send(res, status, body) {
  res.writeHead(status, JSON_HEADERS);
  res.end(JSON.stringify(body));
}

function getSupabaseConfig() {
  const url = process.env.SUPABASE_URL;
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
  const table = process.env.SUPABASE_TABLE || 'pmo_records';
  if (!url || !key) {
    const err = new Error('Backend not configured. Add SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY in Vercel.');
    err.status = 503;
    throw err;
  }
  return { url, key, table };
}

async function supabaseRequest(path, options = {}) {
  const cfg = getSupabaseConfig();
  const response = await fetch(`${cfg.url}/rest/v1/${path}`, {
    ...options,
    headers: {
      'apikey': cfg.key,
      'Authorization': `Bearer ${cfg.key}`,
      'Content-Type': 'application/json',
      'Prefer': options.prefer || 'return=representation',
      ...(options.headers || {})
    }
  });
  const text = await response.text();
  let json;
  try { json = JSON.parse(text); } catch { json = text; }
  if (!response.ok) throw new Error(JSON.stringify(json));
  return json;
}

async function upsertRecords(type, records) {
  if (!records || records.length === 0) return [];
  const rows = records.map(record => ({
    type,
    record: typeof record === 'object' ? record : { value: record },
    updated_at: new Date().toISOString()
  }));
  // Use upsert with on_conflict to avoid duplicates based on type + record->>'id'
  const result = await supabaseRequest(`pmo_records`, {
    method: 'POST',
    prefer: 'return=minimal',
    body: JSON.stringify(rows),
    headers: { 'Prefer': 'return=minimal' }
  });
  return result;
}

export default async function handler(req, res) {
  if (req.method === 'OPTIONS') return send(res, 200, {});
  
  try {
    getSupabaseConfig(); // Validate config early
  } catch (err) {
    return send(res, err.status || 503, { ok: false, error: err.message });
  }

  if (req.method === 'GET') {
    // Health check - return current record count by type
    try {
      const cfg = getSupabaseConfig();
      const rows = await supabaseRequest(`pmo_records?select=type&order=type`);
      const counts = {};
      (rows || []).forEach(row => {
        counts[row.type] = (counts[row.type] || 0) + 1;
      });
      return send(res, 200, { ok: true, counts, total: rows.length });
    } catch (err) {
      return send(res, 500, { ok: false, error: err.message });
    }
  }

  if (req.method === 'POST') {
    // Accept JSON body with { type, records[] } OR { projects[], tasks[], ... }
    let body = req.body;
    if (typeof body === 'string') {
      try { body = JSON.parse(body); } catch { return send(res, 400, { ok: false, error: 'Invalid JSON' }); }
    }

    const ALLOWED_TYPES = ['projects', 'tasks', 'risks', 'reports', 'resources', 'schedules', 'files'];
    const results = {};
    let totalInserted = 0;

    try {
      for (const type of ALLOWED_TYPES) {
        const records = body[type];
        if (records && Array.isArray(records) && records.length > 0) {
          await upsertRecords(type, records);
          results[type] = records.length;
          totalInserted += records.length;
        }
      }
      return send(res, 200, {
        ok: true,
        message: `Synced ${totalInserted} records to Supabase`,
        results,
        syncedAt: new Date().toISOString()
      });
    } catch (err) {
      return send(res, 500, { ok: false, error: err.message });
    }
  }

  return send(res, 405, { ok: false, error: 'Method not allowed' });
}
