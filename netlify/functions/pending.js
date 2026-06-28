// Returns the timesheets waiting for THIS supervisor, grouped one card per
// person's submitted week. Each group: name, site(s), week range, each day's
// work-order lines + hours, the group total, and the row ids to action.
const { getAppToken, validateSupervisorToken, getSupervisorSites, getSubmittedForSites, isSenior, ownsRow, getSitesDetailed, nextPayrollSend } = require('./supervisor');

// Monday-based week start (UTC ISO date) for a given EntryDate.
function weekStartOf(dateStr) {
  const d = new Date(String(dateStr).slice(0, 10) + 'T00:00:00Z');
  if (isNaN(d)) return '';
  const dow = d.getUTCDay();                 // 0=Sun..6=Sat
  const back = (dow + 6) % 7;                 // days since Monday
  d.setUTCDate(d.getUTCDate() - back);
  return d.toISOString().slice(0, 10);
}

exports.handler = async (event) => {
  const headers = { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' };
  const user = await validateSupervisorToken(event.headers && (event.headers.authorization || event.headers.Authorization));
  if (!user) return { statusCode: 401, headers, body: '{"ok":false}' };

  try {
    const token = await getAppToken();
    const sites = await getSupervisorSites(token, user.email);
    if (!sites.length) {
      return { statusCode: 200, headers, body: JSON.stringify({ ok: true, name: user.name, sites: [], groups: [] }) };
    }

    const detailed = await getSitesDetailed(token);
    let nextSend = '';
    for (const s of sites) {
      const info = detailed[s];
      if (info) { const t = nextPayrollSend(info.closeDayIndex); if (!nextSend || t < nextSend) nextSend = t; }
    }

    const rows = await getSubmittedForSites(token, sites);

    // Group by person + Monday-based week. ContractorId is the stable per-person
    // key; fall back to the name if it's missing on a row.
    const groups = {};
    for (const it of rows) {
      const f = it.fields || {};
      if (ownsRow(f, user)) continue;                     // nobody approves their own timesheet
      const person = String(f.ContractorId || f.Title || '').trim() || 'unknown';
      const date = String(f.EntryDate || '').slice(0, 10);
      const wk = weekStartOf(date) || (String(f.BatchID || '').trim() || date);
      const key = person + '|' + wk;
      if (!groups[key]) {
        groups[key] = { key, name: String(f.Title || '').trim() || 'Worker', siteSet: new Set(), weekStart: weekStartOf(date), ids: [], total: 0, byDate: {} };
      }
      const g = groups[key];
      g.ids.push(it.id);
      if (f.Site) g.siteSet.add(String(f.Site).trim());
      const hr = Number(f.Hours) || 0;
      g.total += hr;
      if (!g.weekStart && weekStartOf(date)) g.weekStart = weekStartOf(date);
      if (!g.byDate[date]) g.byDate[date] = [];
      const allow = (String(f.RowType || '').toLowerCase() === 'allowance' || f.Allowance)
        ? (String(f.Allowance || '').trim() || 'Allowance')
        : '';
      g.byDate[date].push({ wo: String(f.WorkOrder || ''), hr, allow, comment: String(f.Comment || '') });
    }

    const out = Object.values(groups).map(g => ({
      key: g.key,
      name: g.name,
      site: [...g.siteSet].sort((a, b) => a.localeCompare(b)).join(', '),
      weekStart: g.weekStart || '',
      total: Math.round(g.total * 100) / 100,
      ids: g.ids,
      days: Object.keys(g.byDate).sort().map(date => ({ date, lines: g.byDate[date] })),
    })).sort((a, b) =>
      (a.name.localeCompare(b.name)) || String(a.weekStart).localeCompare(String(b.weekStart))
    );

    return { statusCode: 200, headers, body: JSON.stringify({ ok: true, name: user.name, sites, groups: out, nextSend }) };
  } catch (err) {
    console.error(err);
    return { statusCode: 502, headers, body: '{"ok":false,"error":"server"}' };
  }
};
