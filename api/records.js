// api/records.js - PMO Live Portal Backend
// Vercel Serverless Function - Supabase CRUD
// SECURITY: Server-side only. No VITE_ prefix. No anon key.

const ALLOWED_TYPES = ['projects','tasks','risks','reports','resources','schedules','files'];

const JSON_HEADERS = {
  'Content-Type': 'application/json',
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET,POST,PUT,DELETE,OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization'
};

function send(res, status, body) {
  res.writeHead(status, JSON_HEADERS);
  res.end(JSON.stringify(body));
}

function isType(value) {
  return ALLOWED_TYPES.includes(value);
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
  const response = await fetch(`${cfg.url}/rest/v1/${cfg.table}${path}`, {
    ...options,
    headers: {
      apikey: cfg.key,
      Authorization: `Bearer ${cfg.key}`,
      'Content-Type': 'application/json',
      Prefer: 'return=representation',
      ...(options.headers || {})
    }
  });
  const text = await response.text();
  let payload = null;
  try { payload = text ? JSON.parse(text) : null; } catch { payload = text; }
  if (!response.ok) {
    const err = new Error(typeof payload === 'string' ? payload : payload?.message || 'Supabase error');
    err.status = response.status;
    err.payload = payload;
    throw err;
  }
  return payload;
}

export default async function handler(req, res) {
  if (req.method === 'OPTIONS') return send(res, 200, { ok: true });

  try {
    const url = new URL(req.url, `https://${req.headers.host || 'localhost'}`);
    const queryType = url.searchParams.get('type');
    const queryId = url.searchParams.get('id');

    // GET
    if (req.method === 'GET') {
      if (queryType && !isType(queryType)) return send(res, 400, { error: 'Invalid record type' });
      const filter = queryType
        ? `?select=id,type,record,created_at,updated_at&type=eq.${encodeURIComponent(queryType)}&order=updated_at.desc`
        : `?select=id,type,record,created_at,updated_at&order=updated_at.desc&limit=500`;
      const rows = await supabaseRequest(filter, { method: 'GET' });
      return send(res, 200, { ok: true, rows: rows || [] });
    }

    // Parse body
    let body = req.body;
    if (!body || typeof body === 'string') {
      try { body = body ? JSON.parse(body) : {}; } catch { body = {}; }
    }

    // POST
    if (req.method === 'POST') {
      if (!isType(body.type)) return send(res, 400, { error: 'Invalid or missing record type' });
      const record = { ...(body.record || {}), source: body.record?.source || 'web-app', createdAt: new Date().toISOString() };
      const rows = await supabaseRequest('', {
        method: 'POST',
        body: JSON.stringify([{ type: body.type, record }])
      });
      return send(res, 201, { ok: true, row: rows?.[0] || null });
    }

    // PUT
    if (req.method === 'PUT') {
      const id = body.id || queryId;
      if (!id) return send(res, 400, { error: 'Missing record id' });
      if (!isType(body.type)) return send(res, 400, { error: 'Invalid or missing record type' });
      const rows = await supabaseRequest(`?id=eq.${encodeURIComponent(id)}`, {
        method: 'PATCH',
        body: JSON.stringify({
          type: body.type,
          record: { ...(body.record || {}), updatedAt: new Date().toISOString() },
          updated_at: new Date().toISOString()
        })
      });
      return send(res, 200, { ok: true, row: rows?.[0] || null });
    }

    // DELETE
    if (req.method === 'DELETE') {
      const id = body.id || queryId;
      if (!id) return send(res, 400, { error: 'Missing record id' });
      await supabaseRequest(`?id=eq.${encodeURIComponent(id)}`, {
        method: 'DELETE',
        headers: { Prefer: 'return=minimal' }
      });
      return send(res, 200, { ok: true, deleted: id });
    }

    return send(res, 405, { error: 'Method not allowed' });

  } catch (error) {
    return send(res, error.status || 500, {
      ok: false,
      error: error.message || 'Unknown backend error',
      details: error.payload || null
    });
  }
}
