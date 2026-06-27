// Approve or reject a group of submitted timesheet rows.
// Body: { ids:[...], action:'approve'|'reject', reason? }
//
// Security: we re-check server-side that EVERY id is currently a Submitted row
// whose Site is one the supervisor manages — so a supervisor can't approve rows
// outside their sites even if they post arbitrary ids.
const { getAppToken, validateSupervisorToken, getSupervisorSites, getSubmittedForSites, updateItem } = require('./supervisor');

exports.handler = async (event) => {
  const headers = { 'Content-Type': 'application/json' };
  const bad = msg => ({ statusCode: 400, headers, body: JSON.stringify({ ok: false, error: msg }) });

  const user = await validateSupervisorToken(event.headers && (event.headers.authorization || event.headers.Authorization));
  if (!user) return { statusCode: 401, headers, body: '{"ok":false}' };

  let data;
  try { data = JSON.parse(event.body || '{}'); } catch { return bad('bad json'); }

  const ids = (Array.isArray(data.ids) ? data.ids : []).map(x => String(x)).filter(Boolean);
  const action = String(data.action || '').toLowerCase();
  const reason = String(data.reason || '').trim().slice(0, 500);
  if (!ids.length) return bad('no ids');
  if (action !== 'approve' && action !== 'reject') return bad('bad action');
  if (action === 'reject' && !reason) return bad('reason required');

  try {
    const token = await getAppToken();
    const sites = await getSupervisorSites(token, user.email);
    if (!sites.length) return { statusCode: 403, headers, body: JSON.stringify({ ok: false, error: 'no sites' }) };

    // Re-fetch what's actually Submitted for this supervisor's sites; only ids
    // present in that set are allowed to be actioned.
    const rows = await getSubmittedForSites(token, sites);
    const allowed = new Map(rows.map(it => [String(it.id), it]));
    const targets = ids.filter(id => allowed.has(id));
    if (!targets.length) return bad('nothing to update');

    const nowIso = new Date().toISOString();
    const when = nowIso.slice(0, 16).replace('T', ' ');                       // e.g. 2026-06-27 14:32
    const who = user.name ? (user.name + ' <' + user.email + '>') : user.email;
    let updated = 0;
    for (const id of targets) {
      const prev = String((allowed.get(id).fields || {}).Notes || '').trim();
      const verb = action === 'approve' ? 'Approved' : 'Rejected';
      const stamp = verb + ' by ' + who + ' on ' + when + (action === 'reject' ? ' — ' + reason : '');
      const notes = prev ? (prev + ' · ' + stamp) : stamp;
      const full = {
        Status: action === 'approve' ? 'Approved' : 'Rejected',
        ApprovedBy: who,                 // who actioned this row (approver, or rejecter)
        ApprovedDate: nowIso,
        Notes: notes,
        ...(action === 'reject' ? { RejectionReason: reason } : {}),
      };
      try {
        await updateItem(token, id, full);
      } catch (e) {
        // The dedicated columns (ApprovedBy/ApprovedDate/RejectionReason) may not
        // exist on the Timesheets list yet — fall back so the decision still lands,
        // with who/when preserved in the Notes field regardless.
        console.warn('decide: full update failed, retrying minimal —', e.message);
        await updateItem(token, id, { Status: full.Status, Notes: notes });
      }
      updated++;
    }

    return { statusCode: 200, headers, body: JSON.stringify({ ok: true, updated }) };
  } catch (err) {
    console.error(err);
    return { statusCode: 502, headers, body: JSON.stringify({ ok: false, error: 'server' }) };
  }
};
