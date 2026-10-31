// Amend + approve a worker's submitted week in ONE step, so a wrong timesheet
// doesn't have to be rejected (which would hold up the pay run). The supervisor can
// change the hours / work order / comment on each line and remove lines; the result
// is saved straight to Approved, stamped as amended.
//
// Body: { ids:[...], edits:{ "<rowId>":{hr,wo,comment} }, removedIds:[...] }
// Security: same re-check as decide.js — every id must currently be a Submitted row
// for a site this supervisor manages, and not the supervisor's own sheet.
//
// Over contract (after the changes): as decide.js — a supervisor's save keeps the week
// Submitted, "awaiting senior approval"; the senior approver's save approves it.
const { getAppToken, validateSupervisorToken, getSupervisorSites, getSubmittedForSites, updateItem, deleteItem, ownsRow, overContract, mayApproveOwn } = require('./supervisor');

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
    const selfOk = await mayApproveOwn(token, user.email);           // testing switch (Settings: SelfApprove)
    const allowed = new Map(rows.filter(it => selfOk || !ownsRow(it.fields, user)).map(it => [String(it.id), it]));
    const targets = ids.filter(id => allowed.has(id));
    if (!targets.length) return bad('nothing to update');

    const nowIso = new Date().toISOString();
    const when = nowIso.slice(0, 16).replace('T', ' ');
    const who = user.name ? (user.name + ' <' + user.email + '>') : user.email;
    const newHours = it => {
      const id = String(it.id);
      if (removed.has(id)) return 0;
      const h = edits[id] && edits[id].hr != null ? parseFloat(edits[id].hr) : NaN;
      return !isNaN(h) && h >= 0 && h <= 24 ? h : (Number((it.fields || {}).Hours) || 0);
    };
    const oc = await overContract(token, user, rows, targets, newHours);
    const hold = oc.overBy > 0 && !oc.senior;
    if (hold && oc.held) return bad('This week is over contract and is waiting for the senior approver.');
    const heldNote = hold ? ` — ${oc.overBy} h over the ${oc.contract} h contract, awaiting senior approval` : '';
    const overNote = !hold && oc.overBy > 0 ? ` (incl. ${oc.overBy} h over the ${oc.contract} h contract)` : '';
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
      const stamp = (hold ? (changed ? 'Amended, Supervisor OK' : 'Supervisor OK') : (changed ? 'Amended & approved' : 'Approved'))
        + ' by ' + who + ' on ' + when + overNote + heldNote + (supNote ? ' · ' + supNote : '');
      const notes = prev ? (prev + ' · ' + stamp) : stamp;

      const full = hold ? { Notes: notes } : { Status: 'Approved', ApprovedBy: who, ApprovedDate: nowIso, Notes: notes };
      if (e.hr != null) { const h = parseFloat(e.hr); if (!isNaN(h) && h >= 0 && h <= 24) full.Hours = h; }
      if (e.wo != null) full.WorkOrder = String(e.wo).trim().slice(0, 200);
      if (e.comment != null) full.Comment = String(e.comment).trim().slice(0, 500);

      try {
        try { await updateItem(token, id, full); }
        catch (err) {
          // A Comment too long for a single-line column (255): keep the first 255 rather than lose it.
          if (!(full.Comment && full.Comment.length > 255 && /^graph update 400\b/.test(String(err.message)))) throw err;
          await updateItem(token, id, { ...full, Comment: full.Comment.slice(0, 255) });
        }
      } catch (err) {
        // Optional columns (ApprovedBy/ApprovedDate/Comment) may not exist — retry
        // keeping the essential changes (hours/work order) + status + a note.
        const minimal = hold ? { Notes: notes } : { Status: 'Approved', Notes: notes };
        if (full.Hours != null) minimal.Hours = full.Hours;
        if (full.WorkOrder != null) minimal.WorkOrder = full.WorkOrder;
        console.warn('amend: full update failed, retrying minimal —', err.message);
        await updateItem(token, id, minimal);
      }
      approved++;
    }

    return { statusCode: 200, headers, body: JSON.stringify({ ok: true, approved: hold ? 0 : approved, held: hold ? approved : 0, overBy: oc.overBy, removed: removedCount }) };
  } catch (err) {
    console.error(err);
    return { statusCode: 502, headers, body: JSON.stringify({ ok: false, error: 'server' }) };
  }
};
