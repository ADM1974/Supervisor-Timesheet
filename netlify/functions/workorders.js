// Let a supervisor view / add / remove WORK ORDERS for THEIR sites, from the app.
// Body: { action:'list' } | { action:'add', site, title } | { action:'remove', id }
// Security: every action is scoped to the sites the signed-in supervisor manages.
// Handles the WorkOrders "Site" column being EITHER a Lookup (writes SiteLookupId,
// resolved from the Sites list) OR plain text (writes the name).
const { getAppToken, validateSupervisorToken, getSupervisorSites, findListId, graphBase } = require('./supervisor');

async function woListId(token) {
  const r = await fetch(`${graphBase()}/lists?$select=id,displayName`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) return null;
  const l = ((await r.json()).value || []).find(x => String(x.displayName || '').toLowerCase().replace(/\s+/g, '') === 'workorders');
  return l ? l.id : null;
}

// The "Site" column on the WorkOrders list: its internal name + whether it's a lookup.
async function siteColumn(token, listId) {
  const r = await fetch(`${graphBase()}/lists/${listId}/columns`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) return { name: 'Site', isLookup: false };
  const cols = (await r.json()).value || [];
  const c = cols.find(x => String(x.displayName || '').trim().toLowerCase() === 'site' || String(x.name || '').trim().toLowerCase() === 'site');
  return c ? { name: c.name, isLookup: !!c.lookup } : { name: 'Site', isLookup: false };
}

// Sites list → { byName: {lowerName: id}, byId: {id: name} } for lookup resolution.
async function sitesMap(token) {
  const id = await findListId(token, 'Sites');
  const out = { byName: {}, byId: {} };
  if (!id) return out;
  const r = await fetch(`${graphBase()}/lists/${id}/items?expand=fields&$top=999`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) return out;
  for (const it of ((await r.json()).value || [])) {
    const nm = String((it.fields || {}).Title || '').trim();
    if (!nm) continue;
    out.byName[nm.toLowerCase()] = it.id;
    out.byId[String(it.id)] = nm;
  }
  return out;
}

exports.handler = async (event) => {
  const headers = { 'Content-Type': 'application/json' };
  const user = await validateSupervisorToken(event.headers && (event.headers.authorization || event.headers.Authorization));
  if (!user) return { statusCode: 401, headers, body: '{"ok":false}' };
  let data; try { data = JSON.parse(event.body || '{}'); } catch { return { statusCode: 400, headers, body: '{"ok":false,"error":"bad json"}' }; }

  try {
    const token = await getAppToken();
    const sites = await getSupervisorSites(token, user.email);
    if (!sites.length) return { statusCode: 403, headers, body: '{"ok":false,"error":"no sites"}' };
    const siteSet = new Set(sites.map(s => s.trim().toLowerCase()));
    const listId = await woListId(token);
    if (!listId) return { statusCode: 400, headers, body: '{"ok":false,"error":"No WorkOrders list found in SharePoint."}' };
    const itemsUrl = `${graphBase()}/lists/${listId}/items`;
    const sc = await siteColumn(token, listId);
    const sm = sc.isLookup ? await sitesMap(token) : null;
    const siteOf = f => sc.isLookup ? (sm.byId[String(f[sc.name + 'LookupId'])] || '') : String(f[sc.name] || '').trim();
    const action = String(data.action || 'list').toLowerCase();

    if (action === 'list') {
      const r = await fetch(`${itemsUrl}?expand=fields&$top=999`, { headers: { Authorization: 'Bearer ' + token } });
      if (!r.ok) throw new Error('graph wo ' + r.status + ' ' + (await r.text()));
      const workOrders = ((await r.json()).value || [])
        .map(it => { const f = it.fields || {}; return { id: it.id, title: String(f.Title || '').trim(), site: siteOf(f) }; })
        .filter(w => w.title && siteSet.has(w.site.toLowerCase()))
        .sort((a, b) => a.site.localeCompare(b.site) || a.title.localeCompare(b.title));
      return { statusCode: 200, headers, body: JSON.stringify({ ok: true, sites, workOrders }) };
    }

    if (action === 'add') {
      const site = String(data.site || '').trim();
      const title = String(data.title || '').trim().slice(0, 200);
      if (!title) return { statusCode: 400, headers, body: '{"ok":false,"error":"Enter a work order"}' };
      if (!siteSet.has(site.toLowerCase())) return { statusCode: 403, headers, body: '{"ok":false,"error":"That isn\'t one of your sites"}' };
      const fields = { Title: title, Active: true };
      if (sc.isLookup) {
        const sid = sm.byName[site.toLowerCase()];
        if (!sid) return { statusCode: 400, headers, body: '{"ok":false,"error":"Could not match the site to the Sites list."}' };
        fields[sc.name + 'LookupId'] = sid;
      } else {
        fields[sc.name] = site;
      }
      const r = await fetch(itemsUrl, { method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields }) });
      if (!r.ok) throw new Error('graph create ' + r.status + ' ' + (await r.text()));
      return { statusCode: 200, headers, body: '{"ok":true}' };
    }

    if (action === 'addmany') {
      const items = Array.isArray(data.items) ? data.items : [];
      let created = 0;
      for (const it of items.slice(0, 200)) {
        const site = String(it.site || '').trim();
        const title = String(it.title || '').trim().slice(0, 200);
        if (!title || !siteSet.has(site.toLowerCase())) continue;
        const fields = { Title: title, Active: true };
        if (sc.isLookup) { const sid = sm.byName[site.toLowerCase()]; if (!sid) continue; fields[sc.name + 'LookupId'] = sid; }
        else fields[sc.name] = site;
        const r = await fetch(itemsUrl, { method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields }) });
        if (r.ok) created++;
      }
      return { statusCode: 200, headers, body: JSON.stringify({ ok: true, created }) };
    }

    if (action === 'remove') {
      const id = String(data.id || '').trim();
      if (!id) return { statusCode: 400, headers, body: '{"ok":false,"error":"no id"}' };
      const gr = await fetch(`${itemsUrl}/${id}?expand=fields`, { headers: { Authorization: 'Bearer ' + token } });
      if (!gr.ok) return { statusCode: 404, headers, body: '{"ok":false,"error":"not found"}' };
      const f = ((await gr.json()).fields) || {};
      if (!siteSet.has(siteOf(f).toLowerCase())) return { statusCode: 403, headers, body: '{"ok":false,"error":"not your site"}' };
      const dr = await fetch(`${itemsUrl}/${id}`, { method: 'DELETE', headers: { Authorization: 'Bearer ' + token } });
      if (!dr.ok && dr.status !== 204) throw new Error('graph delete ' + dr.status + ' ' + (await dr.text()));
      return { statusCode: 200, headers, body: '{"ok":true}' };
    }

    return { statusCode: 400, headers, body: '{"ok":false,"error":"bad action"}' };
  } catch (err) {
    console.error(err);
    return { statusCode: 502, headers, body: JSON.stringify({ ok: false, error: String((err && err.message) || err) }) };
  }
};
