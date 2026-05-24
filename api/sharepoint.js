// api/sharepoint.js - PMO SharePoint Sync Endpoint
// Supports: Microsoft Graph API + Power Automate webhook fallback
// SECURITY: Server-side only. Never expose secrets in frontend.

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

// ─── SharePoint List name mapping ────────────────────────────────────────────
const LIST_MAP = {
  projects:  'PMO Projects',
  tasks:     'PMO Tasks',
  risks:     'PMO Risks',
  reports:   'PMO Status Reports',
  resources: 'PMO Resources',
  schedules: 'PMO Schedules',
  files:     'PMO Files Register'
};

// ─── Get Microsoft Graph token via client_credentials ────────────────────────
async function getGraphToken() {
  const tenantId     = process.env.GRAPH_TENANT_ID;
  const clientId     = process.env.GRAPH_CLIENT_ID;
  const clientSecret = process.env.GRAPH_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    return null; // Graph not configured, will fallback to Power Automate
  }

  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type:    'client_credentials',
    client_id:     clientId,
    client_secret: clientSecret,
    scope:         'https://graph.microsoft.com/.default'
  });

  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString()
  });

  const data = await res.json();
  if (!res.ok) throw new Error(data.error_description || 'Graph token failed');
  return data.access_token;
}

// ─── Create item in SharePoint List via Graph API ────────────────────────────
async function createSharePointItem(token, listName, fields) {
  const siteId  = process.env.SHAREPOINT_SITE_ID;  // e.g. leaderig.sharepoint.com,<site-id>,<web-id>
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${encodeURIComponent(listName)}/items`;

  const res = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ fields })
  });

  const data = await res.json();
  if (!res.ok) throw new Error(data.error?.message || 'SharePoint create failed');
  return data;
}

// ─── Update item in SharePoint List via Graph API ────────────────────────────
async function updateSharePointItem(token, listName, itemId, fields) {
  const siteId = process.env.SHAREPOINT_SITE_ID;
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${encodeURIComponent(listName)}/items/${itemId}/fields`;

  const res = await fetch(url, {
    method: 'PATCH',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(fields)
  });

  const data = await res.json();
  if (!res.ok) throw new Error(data.error?.message || 'SharePoint update failed');
  return data;
}

// ─── Get items from SharePoint List via Graph API ────────────────────────────
async function getSharePointItems(token, listName) {
  const siteId = process.env.SHAREPOINT_SITE_ID;
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${encodeURIComponent(listName)}/items?expand=fields&$top=200`;

  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  });

  const data = await res.json();
  if (!res.ok) throw new Error(data.error?.message || 'SharePoint read failed');
  return data.value || [];
}

// ─── Power Automate webhook fallback ─────────────────────────────────────────
async function sendToPowerAutomate(type, record, action = 'create') {
  const webhookUrl = process.env.POWER_AUTOMATE_WEBHOOK_URL;
  if (!webhookUrl) return { skipped: true, reason: 'POWER_AUTOMATE_WEBHOOK_URL not set' };

  const res = await fetch(webhookUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      action,
      type,
      listName: LIST_MAP[type] || type,
      record,
      timestamp: new Date().toISOString(),
      source: 'pmo-live-portal'
    })
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Power Automate webhook failed: ${text}`);
  }

  return { sent: true, status: res.status };
}

// ─── Map PMO record fields to SharePoint List columns ────────────────────────
function mapToSharePointFields(type, record) {
  const base = {
    Title: record.en || record.title || record.name || record.milestone || `PMO-${type}-${Date.now()}`,
    Source: 'PMO Live Portal',
    SyncedAt: new Date().toISOString()
  };

  switch (type) {
    case 'projects':
      return {
        ...base,
        Title:              record.en || record.title || base.Title,
        ArabicName:         record.ar || '',
        ProjectManager:     record.owner || '',
        Department:         record.department || '',
        Client:             record.client || '',
        Region:             record.region || '',
        Status:             record.status || 'on',
        Priority:           record.priority || 'Medium',
        Progress:           String(record.progress || 0),
        StartDate:          record.start || '',
        DueDate:            record.due || '',
        BudgetUtilization:  String(record.budget || 0),
        TeamUtilization:    String(record.util || 0),
        FilesLink:          record.link || '',
        Notes:              record.notes || ''
      };
    case 'tasks':
      return {
        ...base,
        Title:      record.title || base.Title,
        Project:    record.project || '',
        AssignedTo: record.owner || '',
        Status:     record.status || 'open',
        Priority:   record.priority || 'Medium',
        Progress:   String(record.progress || 0),
        DueDate:    record.due || '',
        Notes:      record.notes || ''
      };
    case 'risks':
      return {
        ...base,
        Title:          record.title || base.Title,
        Project:        record.project || '',
        Owner:          record.owner || '',
        Severity:       record.severity || 'Medium',
        Status:         record.status || 'open',
        Impact:         record.impact || 'Medium',
        Probability:    record.probability || 'Medium',
        TargetDate:     record.due || '',
        MitigationPlan: record.mitigation || ''
      };
    case 'reports':
      return {
        ...base,
        Title:         record.project ? `Report - ${record.project}` : base.Title,
        Project:       record.project || '',
        ReportDate:    record.date || '',
        OverallHealth: record.health || 'on',
        Achievements:  record.achievements || '',
        Challenges:    record.challenges || '',
        NextSteps:     record.nextSteps || '',
        SupportNeeded: record.support || ''
      };
    case 'resources':
      return {
        ...base,
        Title:         record.name || base.Title,
        Employee:      record.name || '',
        Role:          record.role || '',
        Department:    record.department || '',
        Project:       record.project || '',
        CapacityHours: String(record.capacity || 0),
        PlannedHours:  String(record.planned || 0),
        ActualHours:   String(record.actual || 0),
        Utilization:   String(record.util || 0),
        Status:        record.status || 'Available'
      };
    case 'schedules':
      return {
        ...base,
        Title:     record.milestone || base.Title,
        Project:   record.project || '',
        Milestone: record.milestone || '',
        Owner:     record.owner || '',
        StartDate: record.start || '',
        EndDate:   record.due || '',
        Status:    record.status || 'on',
        Progress:  String(record.progress || 0),
        Notes:     record.notes || ''
      };
    case 'files':
      return {
        ...base,
        Title:        record.title || base.Title,
        Project:      record.project || '',
        DocumentType: record.type || 'Other',
        Owner:        record.owner || '',
        UpdatedDate:  record.date || '',
        DocumentLink: record.link || '',
        Notes:        record.notes || ''
      };
    default:
      return base;
  }
}

// ─── Main handler ─────────────────────────────────────────────────────────────
export default async function handler(req, res) {
  if (req.method === 'OPTIONS') return send(res, 200, { ok: true });

  try {
    let body = req.body;
    if (!body || typeof body === 'string') {
      try { body = body ? JSON.parse(body) : {}; } catch { body = {}; }
    }

    const url       = new URL(req.url, `https://${req.headers.host || 'localhost'}`);
    const action    = url.searchParams.get('action') || body.action || 'create';
    const type      = body.type;
    const record    = body.record || {};
    const itemId    = body.sharePointItemId || null;

    if (!type || !LIST_MAP[type]) {
      return send(res, 400, { error: `Invalid or missing type. Must be one of: ${Object.keys(LIST_MAP).join(', ')}` });
    }

    const fields = mapToSharePointFields(type, record);
    let result = {};

    // ── Try Graph API first ──────────────────────────────────────────────────
    const graphConfigured = !!(process.env.GRAPH_TENANT_ID && process.env.GRAPH_CLIENT_ID && process.env.GRAPH_CLIENT_SECRET && process.env.SHAREPOINT_SITE_ID);

    if (graphConfigured) {
      try {
        const token = await getGraphToken();

        if (req.method === 'GET') {
          const items = await getSharePointItems(token, LIST_MAP[type]);
          return send(res, 200, { ok: true, source: 'graph', items });
        }

        if (req.method === 'POST' || action === 'create') {
          const item = await createSharePointItem(token, LIST_MAP[type], fields);
          result = { ok: true, source: 'graph', action: 'created', sharePointId: item.id, fields };
        } else if ((req.method === 'PUT' || action === 'update') && itemId) {
          const item = await updateSharePointItem(token, LIST_MAP[type], itemId, fields);
          result = { ok: true, source: 'graph', action: 'updated', fields };
        } else {
          result = { ok: true, source: 'graph', action: 'skipped', reason: 'No itemId for update' };
        }

        return send(res, 200, result);
      } catch (graphErr) {
        // Graph failed — fallback to Power Automate
        console.error('Graph API failed, trying Power Automate:', graphErr.message);
      }
    }

    // ── Fallback: Power Automate webhook ────────────────────────────────────
    const paResult = await sendToPowerAutomate(type, record, action);
    result = {
      ok: true,
      source: paResult.skipped ? 'none' : 'power-automate',
      action,
      ...paResult,
      fields
    };

    return send(res, 200, result);

  } catch (error) {
    return send(res, error.status || 500, {
      ok: false,
      error: error.message || 'SharePoint sync failed'
    });
  }
}
