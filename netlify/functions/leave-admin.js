// Leave approvals for supervisors. POST { action }:
//   list                       → leave waiting for me at my sites + upcoming approved leave
//   decide { id, action:'approve'|'decline', reason }
//
// Approve: contractors → Approved. Entitlement types (sick, carer's, FDV…) → Acknowledged.
// Staff leave that breaks a rule (see leave-core.js: senior-only type, short notice, long,
// or more than the site's MaxOnLeave off at once — re-checked now) needs the site's senior
// approver as well: a supervisor's Approve makes it "Supervisor OK" and the senior
// approver's Approve makes it Approved. Approved staff leave goes on the person's Teams
// Shifts schedule as "Leave" (the type is never shown on the roster).
const { getAppToken, validateSupervisorToken, getSupervisorSites, getSitesDetailed, isSeniorFor } = require('./supervisor');
const L = require('./leave-core');

const headers = { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' };
const reply = (code, obj) => ({ statusCode: code, headers, body: JSON.stringify(obj) });
const CLASH_RE = /^More than \d+ off at /;

async function emailPerson(token, to, subject, html) {
  const sender = String(process.env.SENDER_EMAIL || '').trim();
  if (!sender || !to) return;
  const r = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(sender)}/sendMail`, {
    method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ message: { subject, body: { contentType: 'HTML', content: html }, toRecipients: [{ emailAddress: { address: to } }] }, saveToSentItems: false }),
  });
  if (!r.ok) console.warn('leave email', r.status, (await r.text()).slice(0, 200));
}
const escH = s => String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const rangeTxt = r => r.start === r.end ? L.fmtDay(r.start) : `${L.fmtDay(r.start)} – ${L.fmtDay(r.end)}`;

exports.handler = async (event) => {
  const user = await validateSupervisorToken(event.headers && (event.headers.authorization || event.headers.Authorization));
  if (!user) return reply(401, { ok: false });
  let data;
  try { data = JSON.parse(event.body || '{}'); } catch (_) { return reply(400, { ok: false, error: 'bad json' }); }
  const me = String(user.email || '').trim().toLowerCase();

  try {
    const token = await getAppToken();
    if (!L.leaveOnFor(await L.leaveSwitch(token), me)) return reply(200, { ok: true, on: false });
    const [sites, detailed, types, leaveSites, { listId, requests }] = await Promise.all([
      getSupervisorSites(token, me), getSitesDetailed(token), L.getLeaveTypes(token), L.getLeaveSites(token), L.getRequests(token),
    ]);
    if (!listId) return reply(200, { ok: true, on: true, setup: 'Create the LeaveRequests list in SharePoint to start (see the README).' });
    const mySite = new Set(sites);
    const typeOf = r => types.find(t => t.title === r.type) || null;
    const senior = site => isSeniorFor(user, [site], detailed);
    const today = L.todayIn('Australia/Sydney');
    const isMine = r => (r.email && r.email === me) || r.personId === user.id;

    if (String(data.action || 'list') === 'list') {
      const waiting = requests.filter(r => mySite.has(r.site) && ['Pending', 'Supervisor OK'].includes(r.status) && !isMine(r));
      const out = waiting.sort((a, b) => a.start.localeCompare(b.start)).map(r => {
        const t = typeOf(r), days = L.daysOf(r);
        const others = L.othersOff(requests, r.site, days, r.id, r.email);
        const busy = days.filter(d => others[d].length).map(d => ({ date: d, people: others[d].map(o => `${o.name} (${L.publicName(typeOf(o))}${o.status === 'Approved' || o.status === 'Acknowledged' ? '' : ', not yet approved'})`) }));
        const held = r.status === 'Supervisor OK';
        return {
          id: r.id, name: r.name, site: r.site, type: r.type, personType: r.personType, start: r.start, end: r.end,
          days: days.length, hours: r.hours, note: r.note, rules: r.rules, evidenceUrl: r.evidenceUrl,
          entitlement: !!(t && t.entitlement), needsAttachment: !!(t && t.needsAttachment), status: r.status,
          held, heldBy: held ? ((/Supervisor OK by ([^\n(]+?)\s*(?:—|\(|$)/m.exec(r.history) || [])[1] || '') : '',
          canAct: !held || senior(r.site), siteMax: (leaveSites[r.site] || {}).max || 0, busy,
        };
      });
      const soon = L.isoOf(L.msOf(today) + 60 * 86400000);
      const upcoming = requests.filter(r => mySite.has(r.site) && L.FINAL_OK.includes(r.status) && r.end >= today && r.start <= soon)
        .sort((a, b) => a.start.localeCompare(b.start))
        .map(r => ({ name: r.name, site: r.site, type: L.publicName(typeOf(r)), start: r.start, end: r.end, days: L.daysOf(r).length }));
      return reply(200, { ok: true, on: true, sites, requests: out, upcoming });
    }

    // ---- decide ----
    const r = requests.find(x => x.id === String(data.id || ''));
    if (!r || !mySite.has(r.site)) return reply(404, { ok: false, error: 'Request not found.' });
    if (isMine(r)) return reply(400, { ok: false, error: 'You can’t approve your own leave.' });
    if (!['Pending', 'Supervisor OK'].includes(r.status)) return reply(400, { ok: false, error: `This request is already ${r.status.toLowerCase()}.` });
    const action = String(data.action || '');
    const t = typeOf(r);
    const who = user.name;

    if (action === 'decline') {
      const reason = String(data.reason || '').trim().slice(0, 500);
      if (!reason) return reply(400, { ok: false, error: 'Give a reason.' });
      await L.updateRequest(token, listId, r.id, { Status: 'Declined', Notified: 'Declined', History: L.addHistory(r, L.stamp(who, 'Declined: ' + reason)) });
      await emailPerson(token, r.email, `Leave declined: ${rangeTxt(r)}`,
        `<p>Hi ${escH(r.name.split(' ')[0])},</p><p>Your ${escH(r.type)} request for <b>${escH(rangeTxt(r))}</b> was declined by ${escH(who)}:</p><blockquote>${escH(reason)}</blockquote><p>Talk to your supervisor if you have questions.</p>`);
      return reply(200, { ok: true, status: 'Declined' });
    }
    if (action !== 'approve') return reply(400, { ok: false, error: 'bad action' });

    let status, rules = r.rules, note;
    if (r.personType === 'Contractor') { status = 'Approved'; rules = []; }
    else if (t && t.entitlement) { status = 'Acknowledged'; rules = []; }
    else {
      // Re-check "too many off" now — other leave may have been approved since it was asked for.
      const days = L.daysOf(r);
      const clashNow = L.evaluateRules({ type: { title: r.type }, start: r.start, days, today: '', site: r.site, siteMax: (leaveSites[r.site] || {}).max || 0, others: L.othersOff(requests, r.site, days, r.id, r.email) });
      rules = [...r.rules.filter(x => !CLASH_RE.test(x)), ...clashNow];
      if (rules.length && !senior(r.site)) {
        if (r.status === 'Supervisor OK') return reply(400, { ok: false, error: 'This is waiting for the senior approver.' });
        await L.updateRequest(token, listId, r.id, { Status: 'Supervisor OK', RulesBroken: rules.join('\n'), History: L.addHistory(r, L.stamp(who, 'Supervisor OK by ' + who + ' — needs senior approval')) });
        return reply(200, { ok: true, status: 'Supervisor OK', held: true });
      }
      status = 'Approved';
      if (rules.length) note = 'Approved by the senior approver despite: ' + rules.join('; ');
    }

    // On the Teams Shifts roster (staff only: contractors aren't in the team).
    let timeOffId = '', shiftsNote = '';
    const teamId = (leaveSites[r.site] || {}).teamId;
    if (r.personType !== 'Contractor' && teamId && /^[0-9a-f-]{36}$/i.test(r.personId)) {
      try { timeOffId = await L.addTimeOff(token, teamId, r.personId, r.start, r.end, r.site); }
      catch (e) { console.warn('leave shifts', e.message); shiftsNote = 'Not added to Teams Shifts: ' + e.message.slice(0, 160); }
    }
    let history = L.addHistory(r, L.stamp(who, (status === 'Acknowledged' ? 'Acknowledged' : 'Approved') + (note ? ' — ' + note : '')));
    if (shiftsNote) history += '\n' + L.stamp('', shiftsNote);
    await L.updateRequest(token, listId, r.id, { Status: status, RulesBroken: rules.join('\n'), History: history, Notified: status, ...(timeOffId ? { ShiftsTimeOffId: timeOffId } : {}) });
    await emailPerson(token, r.email, `Leave ${status.toLowerCase()}: ${rangeTxt(r)}`,
      `<p>Hi ${escH(r.name.split(' ')[0])},</p><p>Your ${escH(r.type)} for <b>${escH(rangeTxt(r))}</b> has been ${status === 'Acknowledged' ? 'recorded' : 'approved'} by ${escH(who)}.</p>`
      + (t && t.needsAttachment && !r.evidenceUrl ? '<p>Please add your medical certificate or statutory declaration in the app (Leave tab).</p>' : ''));
    return reply(200, { ok: true, status, shifts: !!timeOffId, shiftsError: shiftsNote || '' });
  } catch (err) {
    console.error('leave-admin', err);
    if (/not recognized|does not exist|invalidRequest/i.test(String(err.message))) return reply(500, { ok: false, error: 'SharePoint setup: ' + String(err.message).replace(/^.*?"message":"?/, '').slice(0, 160) });
    return reply(502, { ok: false, error: 'Couldn’t reach SharePoint — try again.' });
  }
};
