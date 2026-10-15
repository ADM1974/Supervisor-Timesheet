// Amend + approve a worker's submitted week in ONE step, so a wrong timesheet
// doesn't have to be rejected (which would hold up the pay run). The supervisor can
// change the hours / work order / comment on each line and remove lines; the result
// is saved straight to Approved, stamped as amended.
//
// Body: { ids:[...], edits:{ "<rowId>":{hr,wo,comment} }, removedIds:[...] }
// Security: same re-check as decide.js — every id must currently be a Submitted row
// for a site this supervisor manages, and not the supervisor's own sheet.
const { getAppToken, validateSupervisorToken, getSupervisorSites, getSubmittedForSites, updateItem, deleteItem, ownsRow } = require('./supervisor');

exports.handler = async (event) => {
  const headers = { 'Content-Type': 'application/json' };
  const bad = msg => ({ statusCode: 400, headers, body: JSON.stringify({ ok: false, error: msg }) });

  const user = await validateSupervisorToken(event.headers && (event.headers.authorization || event.headers.Authorization));
  if (!user) return { statusCode: 401, headers, body: '{"ok":false}' };

  let data;
  try { data = JSON.parse(event.body || '{}'); } catch { return bad('bad json'); }

  const ids = (Array.isArray(data.ids) ? data.ids : []).map(x => String(x)).filter(Boolean);
  const edits = (data.edits && typeof data.edits === 'object') ? data.edits : {};
  const removed = new Set((Array.isArray(data.removedIds) ? data.removedIds : []).map(x => String(x)));
  const supNote = String(data.note || '').trim().slice(0, 500);
  if (!ids.length) return bad('no ids');

  try {
    const token = await getAppToken();
    const sites = await getSupervisorSites(token, user.email);
    if (!sites.length) return { statusCode: 403, headers, body: JSON.stringify({ ok: false, error: 'no sites' }) };

    const rows = await getSubmittedForSites(token, sites);
    const allowed = new Map(rows.filter(it => !ownsRow(it.fields, user)).map(it => [String(it.id), it]));
    const targets = ids.filter(id => allowed.has(id));
    if (!targets.length) return bad('nothing to update');

    const nowIso = new Date().toISOString();
    const when = nowIso.slice(0, 16).replace('T', ' ');
    const who = user.name ? (user.name + ' <' + user.email + '>') : user.email;
    let approved = 0, removedCount = 0;

    for (const id of targets) {
      const f = allowed.get(id).fields || {};
      if (removed.has(id)) {
        try { await deleteItem(token, id); removedCount++; } catch (e) { console.warn('amend delete ' + id, e.message); }
        continue;
      }
      const e = edits[id] || {};
      const changed = (e.hr != null) || (e.wo != null) || (e.comment != null);
      const prev = String(f.Notes || '').trim();
      const stamp = (changed ? 'Amended & approved' : 'Approved') + ' by ' + who + ' on ' + when
        + (supNote ? ' — ' + supNote : '');
      const notes = prev ? (prev + ' · ' + stamp) : stamp;

      const full = { Status: 'Approved', ApprovedBy: who, ApprovedDate: nowIso, Notes: notes };
      if (e.hr != null) { const h = parseFloat(e.hr); if (!isNaN(h) && h >= 0 && h <= 24) full.Hours = h; }
      if (e.wo != null) full.WorkOrder = String(e.wo).trim().slice(0, 200);
      if (e.comment != null) full.Comment = String(e.comment).trim().slice(0, 500);

      try {
        await updateItem(token, id, full);
      } catch (err) {
        // Optional columns (ApprovedBy/ApprovedDate/Comment) may not exist — retry
        // keeping the essential changes (hours/work order) + status + a note.
        const minimal = { Status: 'Approved', Notes: notes };
        if (full.Hours != null) minimal.Hours = full.Hours;
        if (full.WorkOrder != null) minimal.WorkOrder = full.WorkOrder;
        console.warn('amend: full update failed, retrying minimal —', err.message);
        await updateItem(token, id, minimal);
      }
      approved++;
    }

    return { statusCode: 200, headers, body: JSON.stringify({ ok: true, approved, removed: removedCount }) };
  } catch (err) {
    console.error(err);
    return { statusCode: 502, headers, body: JSON.stringify({ ok: false, error: 'server' }) };
  }
};
