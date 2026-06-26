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

// The sites this supervisor manages. PRIMARY: the "Manager" People Picker on the
// Sites list — the office just selects supervisors from 365, and we resolve each
// picked person's email from SharePoint's hidden User Information List. FALLBACK:
// a "ManagerEmail" text column, used if a person can't be resolved or is typed in
// directly. A site matches if EITHER route equals the supervisor's email
// (case-insensitive). Returns the site NAMES (Title).
async function getSupervisorSites(token, email) {
  const want = String(email || '').trim().toLowerCase();
  if (!want) return [];
  const sitesListId = await findListId(token, 'Sites');
  if (!sitesListId) throw new Error('Sites list not found');

  const r = await fetch(`${graphBase()}/lists/${sitesListId}/items?expand=fields&$top=999`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) throw new Error('graph sites ' + r.status + ' ' + (await r.text()));
  const rows = (await r.json()).value || [];

  // Best-effort: map SharePoint user LookupId -> email via the User Information
  // List, so a person picked in the Manager column resolves to their email. If
  // that list isn't reachable we silently fall back to the ManagerEmail column.
  const idToEmail = {};
  try {
    const uilId = await findListId(token, 'User Information List');
    if (uilId) {
      const ur = await fetch(`${graphBase()}/lists/${uilId}/items?expand=fields($select=EMail)&$top=999`, { headers: { Authorization: 'Bearer ' + token } });
      if (ur.ok) {
        for (const u of ((await ur.json()).value || [])) {
          const em = String((u.fields || {}).EMail || '').trim().toLowerCase();
          if (em) idToEmail[String(u.id)] = em;
        }
      }
    }
  } catch (e) { /* fall back to ManagerEmail */ }

  return rows.filter(it => {
    const f = it.fields || {};
    const lid = f.ManagerLookupId != null ? String(f.ManagerLookupId) : '';
    const personEmail = lid ? (idToEmail[lid] || '') : '';          // from the People Picker
    const textEmail = String(f.ManagerEmail || '').trim().toLowerCase(); // fallback column
    return personEmail === want || textEmail === want;
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
