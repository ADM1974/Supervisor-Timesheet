// Scheduled (daily 23:00 UTC ≈ 2 hours before the 01:00 UTC payroll send). For each
// site whose report goes out in ~2 hours (its pay week just closed), email ALL its
// approvers a reminder to finish approvals. Schedule is in netlify.toml.
const { getAppToken, getSitesDetailed } = require('./supervisor');

// The NZ day-of-week as it will be at SEND time (01:00 UTC ≈ now + 2h), so this
// matches exactly which sites scheduled-payroll will send.
function sendTimeNzDow() {
  const at = new Date(Date.now() + 2 * 3600 * 1000);
  const nz = new Intl.DateTimeFormat('en-CA', { timeZone: 'Pacific/Auckland', year: 'numeric', month: '2-digit', day: '2-digit' }).format(at);
  const [y, m, d] = nz.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d)).getUTCDay();
}

async function emailReminder(token, toEmails, site) {
  const sender = String(process.env.SENDER_EMAIL || '').trim();
  if (!sender || !toEmails.length) return false;
  const message = {
    subject: `Reminder: ${site} timesheets go to payroll in ~2 hours`,
    body: { contentType: 'HTML', content: `<p>The approved-hours report for <b>${site}</b> will be emailed to payroll (accounts@ael.co) in about <b>2 hours</b>.</p><p>Please make sure every timesheet for ${site} this week is <b>approved</b> before then — anything not approved won't be included.</p>` },
    toRecipients: toEmails.map(a => ({ emailAddress: { address: a } })),
  };
  const r = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(sender)}/sendMail`, {
    method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, body: JSON.stringify({ message, saveToSentItems: false }),
  });
  if (!r.ok) { console.warn('reminder sendMail', r.status, await r.text()); return false; }
  return true;
}

exports.handler = async () => {
  try {
    const token = await getAppToken();
    const detailed = await getSitesDetailed(token);
    const dow = sendTimeNzDow();
    const out = [];
    for (const site of Object.keys(detailed)) {
      const info = detailed[site];
      if (dow !== (info.closeDayIndex + 1) % 7) continue;   // its report sends in ~2h
      const sent = await emailReminder(token, info.approverEmails, site);
      out.push({ site, approvers: info.approverEmails.length, sent });
    }
    console.log('payroll-reminder:', JSON.stringify(out));
    return { statusCode: 200, body: JSON.stringify({ ok: true, out }) };
  } catch (err) {
    console.error('payroll-reminder error', err);
    return { statusCode: 500, body: 'error' };
  }
};
