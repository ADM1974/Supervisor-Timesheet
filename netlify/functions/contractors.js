// Let a supervisor set up CONTRACTORS and their access codes from the app, so a
// new contractor can start using the green app without anyone touching SharePoint.
// Body:
//   { action:'list' }
//   { action:'add', name }                 -> creates a contractor + auto code
//   { action:'regen', id }                 -> issues a fresh access code
//   { action:'setactive', id, active }     -> turn a contractor on/off
//   { action:'rename', id, name }          -> fix a name
// Security: any signed-in supervisor may manage contractors (agreed scope). The
// token is still verified server-side, so only a real supervisor can call this.
const { getAppToken, validateSupervisorToken, getSupervisorSites, findListId, graphBase } = require('./supervisor');

// Unambiguous code: no 0/O/1/I/L so it's easy to read out over the phone.
const ALPHABET = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
function makeCode(len = 6) {
  let s = '';
  for (let i = 0; i < len; i++) s += ALPHABET[Math.floor(Math.random() * ALPHABET.length)];
  return s;
}
// A code not already in use (case-insensitive), trying a few times before giving up.
function uniqueCode(taken) {
  for (let i = 0; i < 40; i++) { const c = makeCode(); if (!taken.has(c.toUpperCase())) return c; }
  return makeCode(8); // extremely unlikely fallback
}

async function readContractors(token, itemsUrl) {
  const r = await fetch(`${itemsUrl}?expand=fields&$top=999`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) throw new Error('graph contractors ' + r.status + ' ' + (await r.text()));
  return ((await r.json()).value || []).map(it => {
    const f = it.fields || {};
    return {
      id: it.id,
      name: String(f.Title || f.Name || '').trim(),
      code: String(f.Code || '').trim(),
      active: f.Active === undefined ? true : f.Active === true, // default-active, like the green app
    };
  });
}

exports.handler = async (event) => {
  const headers = { 'Content-Type': 'application/json' };
  const user = await validateSupervisorToken(event.headers && (event.headers.authorization || event.headers.Authorization));
  if (!user) return { statusCode: 401, headers, body: '{"ok":false}' };
  let data; try { data = JSON.parse(event.body || '{}'); } catch { return { statusCode: 400, headers, body: '{"ok":false,"error":"bad json"}' }; }

  try {
    const token = await getAppToken();
    // Must be a real supervisor (linked to at least one site), but NOT site-scoped.
    const sites = await getSupervisorSites(token, user.email);
    if (!sites.length) return { statusCode: 403, headers, body: '{"ok":false,"error":"no sites"}' };

    const listId = await findListId(token, 'Contractors');
    if (!listId) return { statusCode: 400, headers, body: '{"ok":false,"error":"No Contractors list found in SharePoint."}' };
    const itemsUrl = `${graphBase()}/lists/${listId}/items`;
    const action = String(data.action || 'list').toLowerCase();

    const existing = await readContractors(token, itemsUrl);
    const taken = new Set(existing.map(c => c.code.toUpperCase()).filter(Boolean));

    if (action === 'list') {
      const contractors = existing.slice().sort((a, b) =>
        (a.active === b.active ? 0 : a.active ? -1 : 1) || a.name.localeCompare(b.name));
      return { statusCode: 200, headers, body: JSON.stringify({ ok: true, contractors }) };
    }

    if (action === 'add') {
      const name = String(data.name || '').trim().slice(0, 150);
      if (!name) return { statusCode: 400, headers, body: '{"ok":false,"error":"Enter a name"}' };
      if (existing.some(c => c.name.toLowerCase() === name.toLowerCase()))
        return { statusCode: 400, headers, body: '{"ok":false,"error":"There is already a contractor with that name."}' };
      const code = uniqueCode(taken);
      // Try the full row; retry minimal if a column is missing or named differently.
      const full = { Title: name, Code: code, Active: true };
      let r = await fetch(itemsUrl, { method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields: full }) });
      if (!r.ok) {
        const minimal = { Title: name, Code: code };
        r = await fetch(itemsUrl, { method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields: minimal }) });
      }
      if (!r.ok) throw new Error('graph create ' + r.status + ' ' + (await r.text()));
      return { statusCode: 200, headers, body: JSON.stringify({ ok: true, name, code }) };
    }

    if (action === 'regen') {
      const id = String(data.id || '').trim();
      if (!id) return { statusCode: 400, headers, body: '{"ok":false,"error":"no id"}' };
      if (!existing.some(c => c.id === id)) return { statusCode: 404, headers, body: '{"ok":false,"error":"not found"}' };
      const code = uniqueCode(taken);
      const r = await fetch(`${itemsUrl}/${id}/fields`, { method: 'PATCH', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, body: JSON.stringify({ Code: code }) });
      if (!r.ok) throw new Error('graph regen ' + r.status + ' ' + (await r.text()));
      return { statusCode: 200, headers, body: JSON.stringify({ ok: true, code }) };
    }

    if (action === 'setactive') {
      const id = String(data.id || '').trim();
      if (!id) return { statusCode: 400, headers, body: '{"ok":false,"error":"no id"}' };
      const active = data.active === true;
      const r = await fetch(`${itemsUrl}/${id}/fields`, { method: 'PATCH', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, body: JSON.stringify({ Active: active }) });
      if (!r.ok) throw new Error('graph setactive ' + r.status + ' ' + (await r.text()));
      return { statusCode: 200, headers, body: JSON.stringify({ ok: true, active }) };
    }

    if (action === 'rename') {
      const id = String(data.id || '').trim();
      const name = String(data.name || '').trim().slice(0, 150);
      if (!id || !name) return { statusCode: 400, headers, body: '{"ok":false,"error":"need id + name"}' };
      const r = await fetch(`${itemsUrl}/${id}/fields`, { method: 'PATCH', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, body: JSON.stringify({ Title: name }) });
      if (!r.ok) throw new Error('graph rename ' + r.status + ' ' + (await r.text()));
      return { statusCode: 200, headers, body: JSON.stringify({ ok: true, name }) };
    }

    return { statusCode: 400, headers, body: '{"ok":false,"error":"bad action"}' };
  } catch (err) {
    console.error(err);
    return { statusCode: 502, headers, body: JSON.stringify({ ok: false, error: String((err && err.message) || err) }) };
  }
};
