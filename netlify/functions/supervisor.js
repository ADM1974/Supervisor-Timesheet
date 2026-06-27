// Shared backend for the SUPERVISOR approval app.
//
// Identity: supervisors sign in with Microsoft (MSAL) in the browser; the
// browser sends their ID token, which we VERIFY here against Microsoft
// (signature, tenant, audience) — so identity can't be faked. We then read &
// update the SHARED Timesheets list using the existing APP-ONLY connection
// (same creds as the contractor + staff apps).
//
// This file is also the /supervisor ("me") endpoint, and exports helpers the
// other functions reuse — mirroring the staff app's staff.js pattern.
//
// Environment variables (Netlify):
//   TENANT_ID, CLIENT_ID, CLIENT_SECRET   - app-only creds (SharePoint read/write)
//   SP_SITE_ID, LIST_ID                   - the SharePoint site + Timesheets list
//   SUP_CLIENT_ID                         - the Supervisor sign-in app registration (token audience)

const TOKEN_URL = t => `https://login.microsoftonline.com/${t}/oauth2/v2.0/token`;

// ---- app-only token for SharePoint reads/writes (client credentials) ----
async function getAppToken() {
  const body = new URLSearchParams({
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });
  const r = await fetch(TOKEN_URL(process.env.TENANT_ID), { method: 'POST', body });
  if (!r.ok) throw new Error('token ' + r.status + ' ' + (await r.text()));
  return (await r.json()).access_token;
}

// ---- verify a supervisor's Microsoft ID token, return their identity ----
let _jwks;
async function validateSupervisorToken(authHeader) {
  const m = /^Bearer (.+)$/.exec(authHeader || '');
  if (!m) return null;
  const tenant = process.env.TENANT_ID;
  const { jwtVerify, createRemoteJWKSet } = await import('jose');
  if (!_jwks) _jwks = createRemoteJWKSet(new URL(`https://login.microsoftonline.com/${tenant}/discovery/v2.0/keys`));
  try {
    const { payload } = await jwtVerify(m[1], _jwks, {
      issuer: `https://login.microsoftonline.com/${tenant}/v2.0`,
      audience: process.env.SUP_CLIENT_ID,
    });
    const id = String(payload.oid || payload.sub || '').trim();
    if (!id) return null;
    return {
      id,                                                        // stable per-user key
      name: String(payload.name || '').trim() || 'Supervisor',
      email: String(payload.preferred_username || payload.email || '').trim(),
    };
  } catch (e) { return null; }
}

// ===== SharePoint helpers =====
const graphBase = () => `https://graph.microsoft.com/v1.0/sites/${process.env.SP_SITE_ID}`;
const itemsUrl = () => `${graphBase()}/lists/${process.env.LIST_ID}/items`;

// Find a list id by its displayName (case-insensitive). Mirrors staff.js getActiveSites.
async function findListId(token, name) {
  const r = await fetch(`${graphBase()}/lists?$select=id,displayName`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) throw new Error('graph lists ' + r.status + ' ' + (await r.text()));
  const list = ((await r.json()).value || []).find(l => String(l.displayName || '').toLowerCase() === String(name).toLowerCase());
  return list ? list.id : null;
}

// Turns a manager's name into the company email — format: first-initial.surname@ael.co
// (e.g. "Andrew Corner" → a.corner@ael.co). Returns '' if it can't form one.
function emailFromName(name) {
  const parts = String(name || '').trim().split(/\s+/).filter(Boolean);
  if (parts.length < 2) return '';
  const first = parts[0].replace(/[^a-z]/gi, '').charAt(0).toLowerCase();
  const last = parts[parts.length - 1].replace(/[^a-z0-9-]/gi, '').toLowerCase();
  return (first && last) ? `${first}.${last}@ael.co` : '';
}

// The sites this supervisor manages. A site can have MULTIPLE approvers (for cover):
// list names in "ManagerName" separated by ';' (each → email via the company format),
// and/or explicit emails in "ManagerEmail" (also ';'-separated). A site matches if the
// signed-in supervisor's email is ANY of those. Returns the site NAMES (Title).
async function getSupervisorSites(token, email) {
  const want = String(email || '').trim().toLowerCase();
  if (!want) return [];
  const sitesListId = await findListId(token, 'Sites');
  if (!sitesListId) throw new Error('Sites list not found');

  const r = await fetch(`${graphBase()}/lists/${sitesListId}/items?expand=fields&$top=999`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) throw new Error('graph sites ' + r.status + ' ' + (await r.text()));
  const rows = (await r.json()).value || [];

  const split = s => String(s || '').split(/[;\n]/).map(x => x.trim()).filter(Boolean);
  return rows.filter(it => {
    const f = it.fields || {};
    const emails = new Set();
    for (const e of split(f.ManagerEmail)) emails.add(e.toLowerCase());       // explicit override emails
    for (const n of split(f.Manager || f.ManagerName)) {                      // names → generated emails
      const g = emailFromName(n); if (g) emails.add(g);
    }
    return emails.has(want);
  })
    .map(it => String((it.fields || {}).Title || '').trim())
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));
}

// All Submitted Timesheets rows whose Site is one of the supervisor's sites.
// We $filter on Status (an indexed-or-not column) using the Prefer header, then
// filter Site in code (Site is a choice; safest done locally).
async function getSubmittedForSites(token, siteNames) {
  const wanted = new Set((siteNames || []).map(s => String(s).trim()));
  if (!wanted.size) return [];
  const url = `${itemsUrl()}?expand=fields&$top=999&$filter=fields/Status eq 'Submitted'`;
  const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token, Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' } });
  if (!r.ok) throw new Error('graph submitted ' + r.status + ' ' + (await r.text()));
  return ((await r.json()).value || [])
    .filter(it => wanted.has(String((it.fields || {}).Site || '').trim()));
}

async function updateItem(token, id, fields) {
  const r = await fetch(`${itemsUrl()}/${id}/fields`, {
    method: 'PATCH',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify(fields),
  });
  if (!r.ok) throw new Error('graph update ' + r.status + ' ' + (await r.text()));
}

// ===== payroll-reporting helpers =====
const WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

// Sites list → { name: { name, closeDayIndex, approverEmails:[...] } }. Used to know
// each site's pay-week close day + who to cc (its approvers' emails).
async function getSitesDetailed(token) {
  const id = await findListId(token, 'Sites');
  if (!id) throw new Error('Sites list not found');
  const r = await fetch(`${graphBase()}/lists/${id}/items?expand=fields&$top=999`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) throw new Error('graph sites ' + r.status + ' ' + (await r.text()));
  const split = s => String(s || '').split(/[;\n]/).map(x => x.trim()).filter(Boolean);
  const out = {};
  for (const it of ((await r.json()).value || [])) {
    const f = it.fields || {};
    const name = String(f.Title || '').trim();
    if (!name || f.Active === false) continue;
    const emails = new Set();
    for (const e of split(f.ManagerEmail)) emails.add(e.toLowerCase());
    for (const n of split(f.Manager || f.ManagerName)) { const g = emailFromName(n); if (g) emails.add(g); }
    const idx = WEEKDAYS.findIndex(d => d.toLowerCase() === String(f.CloseDay || '').trim().toLowerCase());
    out[name] = { name, closeDayIndex: idx < 0 ? 0 : idx, approverEmails: [...emails] };
  }
  return out;
}

// Approved Timesheets rows for the given site names within [weekStart,weekEnd] (ISO dates).
async function getApprovedForSites(token, siteNames, weekStart, weekEnd) {
  const wanted = new Set((siteNames || []).map(s => String(s).trim()));
  if (!wanted.size) return [];
  const url = `${itemsUrl()}?expand=fields&$top=999&$filter=fields/Status eq 'Approved'`;
  const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token, Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' } });
  if (!r.ok) throw new Error('graph approved ' + r.status + ' ' + (await r.text()));
  return ((await r.json()).value || []).filter(it => {
    const f = it.fields || {};
    if (!wanted.has(String(f.Site || '').trim())) return false;
    const ed = String(f.EntryDate || '').slice(0, 10);
    return ed >= weekStart && ed <= weekEnd;
  });
}

// Today in NZ + its day-of-week (0=Sun..6=Sat).
function nzToday() {
  const s = new Intl.DateTimeFormat('en-CA', { timeZone: 'Pacific/Auckland', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
  const [y, m, d] = s.split('-').map(Number);
  return { y, m, d, dow: new Date(Date.UTC(y, m - 1, d)).getUTCDay() };
}
const _isoDay = ms => new Date(ms).toISOString().slice(0, 10);

// The 7-day pay week ENDING on closeDayIndex. weeksBack: 0 = current open week,
// 1 = the most recently completed week (what payroll pays), etc.
function weekForClose(closeDayIndex, weeksBack) {
  const { y, m, d, dow } = nzToday();
  const today = Date.UTC(y, m - 1, d);
  let end = today + ((closeDayIndex - dow + 7) % 7) * 86400000;   // this/next close day = current open week end
  end -= (weeksBack || 0) * 7 * 86400000;
  return { weekStart: _isoDay(end - 6 * 86400000), weekEnd: _isoDay(end) };
}

// ---- /supervisor endpoint: confirm sign-in and return the supervisor's name ----
exports.handler = async (event) => {
  const headers = { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' };
  const user = await validateSupervisorToken(event.headers && (event.headers.authorization || event.headers.Authorization));
  if (!user) return { statusCode: 401, headers, body: '{"ok":false}' };
  return { statusCode: 200, headers, body: JSON.stringify({ ok: true, name: user.name, email: user.email }) };
};

exports.getAppToken = getAppToken;
exports.validateSupervisorToken = validateSupervisorToken;
exports.findListId = findListId;
exports.getSupervisorSites = getSupervisorSites;
exports.getSubmittedForSites = getSubmittedForSites;
exports.updateItem = updateItem;
exports.itemsUrl = itemsUrl;
exports.graphBase = graphBase;
exports.getSitesDetailed = getSitesDetailed;
exports.getApprovedForSites = getApprovedForSites;
exports.nzToday = nzToday;
exports.weekForClose = weekForClose;
