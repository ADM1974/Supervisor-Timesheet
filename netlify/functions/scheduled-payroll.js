// Scheduled (daily) payroll send. For each site whose pay week CLOSED the previous
// day (today, NZ, is the day after that site's CloseDay), email payroll the
// approved-hours report for that just-completed week, cc the site's supervisor.
// The cron schedule is set in netlify.toml ([functions."scheduled-payroll"]).
//
// Late approvals after this runs are covered by the manual "send now" button in
// the app, which re-sends the same week.
const { getAppToken, getSitesDetailed, nzToday, weekForClose } = require('./supervisor');
const { buildAndSendForSite } = require('./report');

exports.handler = async () => {
  try {
    const token = await getAppToken();
    const detailed = await getSitesDetailed(token);
    const { dow } = nzToday();
    const out = [];
    for (const site of Object.keys(detailed)) {
      const info = detailed[site];
      if (dow !== (info.closeDayIndex + 1) % 7) continue;   // only the day after THIS site's close day
      const { weekStart, weekEnd } = weekForClose(info.closeDayIndex, 1);
      try { out.push(await buildAndSendForSite(token, site, weekStart, weekEnd, info.approverEmails)); }
      catch (e) { out.push({ site, sent: false, error: String((e && e.message) || e) }); }
    }
    console.log('scheduled-payroll:', JSON.stringify(out));
    return { statusCode: 200, body: JSON.stringify({ ok: true, out }) };
  } catch (err) {
    console.error('scheduled-payroll error', err);
    return { statusCode: 500, body: 'error' };
  }
};
