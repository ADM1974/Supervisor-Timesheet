// Let a supervisor view / add / remove WORK ORDERS for THEIR sites, from the app.
// Body: { action:'list' } | { action:'add', site, title } | { action:'addmany', items:[{site,title}] }
//       | { action:'remove', id, site }
// Security: every action is scoped to the sites the signed-in supervisor manages, plus any
// site whose Sites.WorkOrderEditors names them (work orders only, no approvals).
// Handles the WorkOrders "Site" column being a Lookup (writes SiteLookupId, resolved from
// the Sites list), a multi-value Lookup ("Allow multiple values"), or plain text ("BNE; SYD").
// One work order can be at several sites: adding a title that already exists adds the site
// to that row, and removing it from a site only deletes the row once no site is left
// (a single-value Lookup column can't hold two sites, so there it's one row per site).
const { getAppToken, validateSupervisorToken, getSupervisorSites, getWorkOrderEditorSites, findListId, graphBase } = require('./supervisor');

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
  return c ? { name: c.name, isLookup: !!c.lookup, multi: !!(c.lookup && c.lookup.allowMultipleValues) } : { name: 'Site', isLookup: false, multi: false };
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
    // The supervisor's own sites, plus any where they're named in Sites.WorkOrderEditors.
    const [supSites, edSites] = await Promise.all([getSupervisorSites(token, user.email), getWorkOrderEditorSites(token, user.email)]);
    const sites = [...new Set([...supSites, ...edSites])].sort((a, b) => a.localeCompare(b));
    if (!sites.length) return { statusCode: 403, headers, body: '{"ok":false,"error":"no sites"}' };
    const siteSet = new Set(sites.map(s => s.trim().toLowerCase()));
    const listId = await woListId(token);
    if (!listId) return { statusCode: 400, headers, body: '{"ok":false,"error":"No WorkOrders list found in SharePoint."}' };
    const itemsUrl = `${graphBase()}/lists/${listId}/items`;
    const sc = await siteColumn(token, listId);
    const sm = sc.isLookup ? await sitesMap(token) : null;
    const canShare = !sc.isLookup || sc.multi;   // can one row hold several sites?
    // A row's sites → [names]
    const sitesOf = f => {
      if (!sc.isLookup) return String(f[sc.name] || '').split(/[;,]/).map(x => x.trim()).filter(Boolean);
      const ids = [];
      const idv = f[sc.name + 'LookupId'];
      (Array.isArray(idv) ? idv : [idv]).forEach(v => { if (v != null && v !== '') ids.push(String(v)); });
      if (Array.isArray(f[sc.name])) f[sc.name].forEach(x => { if (x && x.LookupId != null) ids.push(String(x.LookupId)); });
      return [...new Set(ids)].map(id => sm.byId[id]).filter(Boolean);
    };
    // [names] → the fields to write (null if a name isn't in the Sites list)
    const siteFields = names => {
      if (!sc.isLookup) return { [sc.name]: names.join('; ') };
      const ids = names.map(n => sm.byName[n.toLowerCase()]);
      if (ids.some(x => !x)) return null;
      return sc.multi ? { [sc.name + 'LookupId@odata.type']: 'Collection(Edm.Int32)', [sc.name + 'LookupId']: ids.map(Number) } : { [sc.name + 'LookupId']: ids[0] };
    };
    const send = async (url, method, fields) => {
      const r = await fetch(url, { method, headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, body: JSON.stringify(method === 'POST' ? { fields } : fields) });
      if (!r.ok) throw new Error('graph ' + method + ' ' + r.status + ' ' + (await r.text()));
      return r.json().catch(() => ({}));
    };
    const allRows = async () => {
      const r = await fetch(`${itemsUrl}?expand=fields&$top=999`, { headers: { Authorization: 'Bearer ' + token } });
      if (!r.ok) throw new Error('graph wo ' + r.status + ' ' + (await r.text()));
      return ((await r.json()).value || []).map(it => { const f = it.fields || {}; return { id: it.id, title: String(f.Title || '').trim(), sites: sitesOf(f) }; });
    };
    // Add a title at a site: onto an existing row with that title when the column can share,
    // else a new row. → true if it changed something.
    const addAt = async (rows, site, title) => {
      const key = title.toLowerCase();
      const same = rows.filter(w => w.title.toLowerCase() === key);
      if (same.some(w => w.sites.some(s => s.toLowerCase() === site.toLowerCase()))) return false;   // already there
      const row = canShare && same[0];
      const names = row ? [...row.sites, site] : [site];
      const sf = siteFields(names);
      if (!sf) return false;
      if (row) { await send(`${itemsUrl}/${row.id}/fields`, 'PATCH', sf); row.sites = names; }
      else { const made = await send(itemsUrl, 'POST', { Title: title, Active: true, ...sf }); rows.push({ id: made.id, title, sites: names }); }
      return true;
    };
    const action = String(data.action || 'list').toLowerCase();

    if (action === 'list') {
      const workOrders = (await allRows())
        .filter(w => w.title && w.sites.some(s => siteSet.has(s.toLowerCase())))
        .sort((a, b) => a.title.localeCompare(b.title));
      return { statusCode: 200, headers, body: JSON.stringify({ ok: true, sites, workOrders, multiSite: canShare }) };
    }

    if (action === 'add') {
      const site = String(data.site || '').trim();
      const title = String(data.title || '').trim().slice(0, 200);
      if (!title) return { statusCode: 400, headers, body: '{"ok":false,"error":"Enter a work order"}' };
      if (!siteSet.has(site.toLowerCase())) return { statusCode: 403, headers, body: '{"ok":false,"error":"That isn\'t one of your sites"}' };
      if (!siteFields([site])) return { statusCode: 400, headers, body: '{"ok":false,"error":"Could not match the site to the Sites list."}' };
      await addAt(await allRows(), site, title);
      return { statusCode: 200, headers, body: '{"ok":true}' };
    }

    if (action === 'addmany') {
      const items = Array.isArray(data.items) ? data.items : [];
      const rows = await allRows();
      let created = 0;
      for (const it of items.slice(0, 200)) {
        const site = String(it.site || '').trim();
        const title = String(it.title || '').trim().slice(0, 200);
        if (!title || !siteSet.has(site.toLowerCase())) continue;
        try { if (await addAt(rows, site, title)) created++; } catch (e) { console.warn('wo add', title, e.message); }
      }
      return { statusCode: 200, headers, body: JSON.stringify({ ok: true, created }) };
    }

    if (action === 'remove') {
      const id = String(data.id || '').trim();
      if (!id) return { statusCode: 400, headers, body: '{"ok":false,"error":"no id"}' };
      const gr = await fetch(`${itemsUrl}/${id}?expand=fields`, { headers: { Authorization: 'Bearer ' + token } });
      if (!gr.ok) return { statusCode: 404, headers, body: '{"ok":false,"error":"not found"}' };
      const f = ((await gr.json()).fields) || {};
      const site = String(data.site || '').trim().toLowerCase();
      const has = sitesOf(f);
      // take off the given site (or, from an older app, every site of this supervisor's)
      const drop = s => site ? s.toLowerCase() === site && siteSet.has(site) : siteSet.has(s.toLowerCase());
      if (!has.some(drop)) return { statusCode: 403, headers, body: '{"ok":false,"error":"not your site"}' };
      const left = has.filter(s => !drop(s));
      const sf = left.length && siteFields(left);
      if (sf) { await send(`${itemsUrl}/${id}/fields`, 'PATCH', sf); return { statusCode: 200, headers, body: '{"ok":true,"kept":true}' }; }
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
