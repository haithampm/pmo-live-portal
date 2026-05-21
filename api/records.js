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
  return ['projects', 'tasks', 'risks', 'reports', 'resources', 'schedules', 'files'].includes(value);
}

function getSupabaseConfig() {
  return {
    url: process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL,
    key: process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_ANON_KEY || process.env.VITE_SUPABASE_ANON_KEY,
    table: process.env.SUPABASE_TABLE || 'pmo_records'
  };
}

async function supabaseRequest(path, options = {}) {
  const cfg = getSupabaseConfig();
  if (!cfg.url || !cfg.key) {
    const err = new Error('Backend is not configured. Add SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY in Vercel.');
    err.status = 503;
    throw err;
  }
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
    const err = new Error(typeof payload === 'string' ? payload : payload?.message || 'Supabase request failed');
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
    const type = url.searchParams.get('type') || req.body?.type;

    if (req.method === 'GET') {
      if (type && !isType(type)) return send(res, 400, { error: 'Invalid record type' });
      const query = type
        ? `?select=id,type,record,created_at,updated_at&type=eq.${encodeURIComponent(type)}&order=updated_at.desc`
        : '?select=id,type,record,created_at,updated_at&order=updated_at.desc';
      const rows = await supabaseRequest(query, { method: 'GET' });
      return send(res, 200, { ok: true, rows });
    }

    let body = req.body;
    if (!body || typeof body === 'string') {
      body = body ? JSON.parse(body) : {};
    }

    if (!isType(body.type)) return send(res, 400, { error: 'Invalid or missing record type' });

    if (req.method === 'POST') {
      const record = { ...(body.record || {}), source: body.record?.source || 'web-app' };
      const rows = await supabaseRequest('', {
        method: 'POST',
        body: JSON.stringify([{ type: body.type, record }])
      });
      return send(res, 201, { ok: true, row: rows?.[0] || null });
    }

    if (req.method === 'PUT') {
      if (!body.id) return send(res, 400, { error: 'Missing record id' });
      const rows = await supabaseRequest(`?id=eq.${encodeURIComponent(body.id)}`, {
        method: 'PATCH',
        body: JSON.stringify({ record: body.record || {}, updated_at: new Date().toISOString() })
      });
      return send(res, 200, { ok: true, row: rows?.[0] || null });
    }

    if (req.method === 'DELETE') {
      if (!body.id) return send(res, 400, { error: 'Missing record id' });
      const rows = await supabaseRequest(`?id=eq.${encodeURIComponent(body.id)}`, { method: 'DELETE' });
      return send(res, 200, { ok: true, row: rows?.[0] || null });
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
