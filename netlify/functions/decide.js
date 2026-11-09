// Approve or reject a group of submitted timesheet rows.
// Body: { ids:[...], action:'approve'|'reject', reason?, rejectDates? }
// A reject can send back only some days (rejectDates): the rest of the card is approved in
// the same step, so those days still get paid. No rejectDates = the whole card goes back.
// The worker's app then locks the approved days, so fixing the rest never touches them.
//
// Security: we re-check server-side that EVERY id is currently a Submitted row
// whose Site is one the supervisor manages — so a supervisor can't approve rows
// outside their sites even if they post arbitrary ids.
//
// Over contract: a week with more hours than the worker's contract (Workers.ContractHours)
// needs the site's senior approver too. A supervisor's Approve then only marks it
// "awaiting senior approval" (still Submitted); the senior approver's Approve approves it.
//
// Per week: rows of a pay week that hasn't reached its last day yet are never approved. Approve is
// refused for them; a Reject sends back the ticked days and leaves the other days waiting.
const { getAppToken, validateSupervisorToken, getSupervisorSites, getSubmittedForSites, updateItem, ownsRow, overContract, mayApproveOwn, weekRunning, payWeekEndOf } = require('./supervisor');

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
    const selfOk = await mayApproveOwn(token, user.email);           // testing switch (Settings: SelfApprove)
    const allowed = new Map(rows.filter(it => selfOk || !ownsRow(it.fields, user)).map(it => [String(it.id), it]));
    const targets = ids.filter(id => allowed.has(id));
    if (!targets.length) return bad('nothing to update');

    const nowIso = new Date().toISOString();
    const when = nowIso.slice(0, 16).replace('T', ' ');                       // e.g. 2026-06-27 14:32
    const who = user.name ? (user.name + ' <' + user.email + '>') : user.email;

    // Which rows go back and which are approved
    const dateOf = id => String((allowed.get(id).fields || {}).EntryDate || '').slice(0, 10);
    const pick = action === 'reject' && Array.isArray(data.rejectDates) && data.rejectDates.length
      ? new Set(data.rejectDates.map(d => String(d).slice(0, 10))) : null;
    const rejectSet = new Set(action === 'reject' ? targets.filter(id => !pick || pick.has(dateOf(id))) : []);
    if (action === 'reject' && !rejectSet.size) return bad('Tick at least one day to send back.');
    const running = id => weekRunning(allowed.get(id).fields || {});
    if (action === 'approve' && targets.some(running)) {
      const we = payWeekEndOf(allowed.get(targets.find(running)).fields || {});
      const day = new Date(we + 'T00:00:00Z').toLocaleDateString('en-AU', { weekday: 'long', day: 'numeric', month: 'short', timeZone: 'UTC' });
      return bad(`This week is still running – approve it from ${day}, its last day. To fix a day now, use Reject and tick that day.`);
    }
    const approveIds = targets.filter(id => !rejectSet.has(id) && !running(id));   // a running week's other days stay waiting
    const kept = targets.filter(id => !rejectSet.has(id) && running(id)).length;

    const decideRow = async (id, status, stamp, extra) => {
      const prev = String((allowed.get(id).fields || {}).Notes || '').trim();
      const notes = prev ? (prev + ' · ' + stamp) : stamp;
      try {
        await updateItem(token, id, { Status: status, ApprovedBy: who, ApprovedDate: nowIso, Notes: notes, ...(extra || {}) });   // ApprovedBy = who actioned it (approver, or rejecter)
      } catch (e) {
        // The dedicated columns (ApprovedBy/ApprovedDate/RejectionReason) may not
        // exist on the Timesheets list yet — fall back so the decision still lands,
        // with who/when preserved in the Notes field regardless.
        console.warn('decide: full update failed, retrying minimal —', e.message);
        await updateItem(token, id, { Status: status, Notes: notes });
      }
    };

    let rejected = 0;
    for (const id of rejectSet) { await decideRow(id, 'Rejected', `Rejected by ${who} on ${when} — ${reason}`, { RejectionReason: reason }); rejected++; }

    // Approve the rest. Hours sent back don't count towards the contract check.
    let updated = 0;
    if (approveIds.length) {
      const oc = await overContract(token, user, rows, approveIds, it => rejectSet.has(String(it.id)) ? 0 : (Number((it.fields || {}).Hours) || 0));
      if (oc.overBy > 0 && !oc.senior) {
        if (oc.held) {
          if (!rejected) return bad('This week is over contract and is waiting for the senior approver.');
          return { statusCode: 200, headers, body: JSON.stringify({ ok: true, updated: 0, rejected }) };   // the rest stays with the senior approver
        }
        const stamp = `Supervisor OK by ${who} on ${when} — ${oc.overBy} h over the ${oc.contract} h contract, awaiting senior approval`;
        for (const id of approveIds) {
          const prev = String((allowed.get(id).fields || {}).Notes || '').trim();
          await updateItem(token, id, { Notes: prev ? (prev + ' · ' + stamp) : stamp });
        }
        return { statusCode: 200, headers, body: JSON.stringify({ ok: true, updated: 0, rejected, held: approveIds.length, overBy: oc.overBy }) };
      }
      const overNote = oc.overBy > 0 ? ` (incl. ${oc.overBy} h over the ${oc.contract} h contract)` : '';
      for (const id of approveIds) { await decideRow(id, 'Approved', `Approved by ${who} on ${when}${overNote}`); updated++; }
    }

    return { statusCode: 200, headers, body: JSON.stringify({ ok: true, updated, rejected, kept }) };
  } catch (err) {
    console.error(err);
    return { statusCode: 502, headers, body: JSON.stringify({ ok: false, error: 'server' }) };
  }
};
