// Scheduled (every 10 minutes): emails approvers about leave waiting for them.
//   Pending       → the site's supervisors (and senior approvers)
//   Supervisor OK → the site's senior approver(s) (or env SENIOR_APPROVERS)
// Each request is emailed once per stage (LeaveRequests.Notified holds the stage sent).
// Confidential leave types are named only as "leave" in the email. Does nothing while
// leave is switched off (see leave-core.js).
const { getAppToken, getSitesDetailed } = require('./supervisor');
const L = require('./leave-core');

const escH = s => String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

async function send(token, to, subject, html) {
  const sender = String(process.env.SENDER_EMAIL || '').trim();
  if (!sender || !to.length) return false;
  const r = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(sender)}/sendMail`, {
    method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ message: { subject, body: { contentType: 'HTML', content: html }, toRecipients: to.map(a => ({ emailAddress: { address: a } })) }, saveToSentItems: false }),
  });
  if (!r.ok) { console.warn('leave-notify sendMail', r.status, (await r.text()).slice(0, 200)); return false; }
  return true;
}

exports.handler = async () => {
  const out = [];
  try {
    const token = await getAppToken();
    if (!(await L.leaveSwitch(token)).on) return { statusCode: 200, body: 'leave off' };
    const [{ listId, requests }, detailed, types] = await Promise.all([L.getRequests(token), getSitesDetailed(token), L.getLeaveTypes(token)]);
    if (!listId) return { statusCode: 200, body: 'no list' };
    const envSeniors = String(process.env.SENIOR_APPROVERS || '').split(/[;,\n]/).map(s => s.trim().toLowerCase()).filter(Boolean);
    const app = process.env.URL || '';
    for (const r of requests) {
      if (!['Pending', 'Supervisor OK'].includes(r.status) || r.notified === r.status) continue;
      const site = detailed[r.site] || {};
      let to = r.status === 'Pending' ? (site.approverEmails || []) : ((site.seniorEmails || []).length ? site.seniorEmails : envSeniors);
      to = [...new Set(to)].filter(e => e && e !== r.email);
      const t = types.find(x => x.title === r.type);
      const what = t && t.confidential ? 'leave' : r.type;
      const when = r.start === r.end ? L.fmtDay(r.start) : `${L.fmtDay(r.start)} – ${L.fmtDay(r.end)}`;
      const senior = r.status === 'Supervisor OK';
      const html = `<p>${escH(r.name)} (${escH(r.site)}) has asked for <b>${escH(what)}</b>: <b>${escH(when)}</b> (${L.daysOf(r).length} working day${L.daysOf(r).length === 1 ? '' : 's'}).</p>`
        + (senior ? '<p>Their supervisor has OK’d it, but it needs <b>your senior approval</b>:</p>' : '')
        + (r.rules.length ? `<ul>${r.rules.map(x => `<li>${escH(x)}</li>`).join('')}</ul>` : '')
        + `<p><a href="${escH(app)}">Open the Approvals app</a> to approve or decline.</p>`;
      const ok = await send(token, to, `${senior ? 'Senior approval needed' : 'Leave request'}: ${r.name}, ${when}`, html);
      if (ok || !to.length) await L.updateRequest(token, listId, r.id, { Notified: r.status });
      out.push(`${r.name} ${r.status}: ${ok ? 'emailed ' + to.length : 'not sent'}`);
    }
  } catch (e) {
    out.push('failed: ' + e.message);
  }
  console.log('leave-notify', out.join(' | ') || 'nothing to send');
  return { statusCode: 200, body: out.join('\n') };
};
