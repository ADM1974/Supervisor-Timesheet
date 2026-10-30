// Returns the timesheets waiting for THIS supervisor, grouped one card per
// person's submitted week. Each group: name, site(s), week range, each day's
// work-order lines + hours, the group total, and the row ids to action.
// Weeks over the worker's contract hours carry contract / overBy, and once the supervisor
// has OK'd them, held = true until the site's senior approver approves (canSenior).
const { getAppToken, validateSupervisorToken, getSupervisorSites, getSubmittedForSites, ownsRow, getSitesDetailed, nextPayrollSend, weekStartOf, personWeekKey, jobAllowancesOf, getContractHours, isSeniorFor, HELD_RE } = require('./supervisor');

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

    const [rows, contracts] = await Promise.all([getSubmittedForSites(token, sites), getContractHours(token)]);

    // Group by person + Monday-based week. ContractorId is the stable per-person
    // key; fall back to the name if it's missing on a row.
    const groups = {};
    for (const it of rows) {
      const f = it.fields || {};
      if (ownsRow(f, user)) continue;                     // nobody approves their own timesheet
      const date = String(f.EntryDate || '').slice(0, 10);
      const key = personWeekKey(f);
      if (!groups[key]) {
        groups[key] = { key, name: String(f.Title || '').trim() || 'Worker', siteSet: new Set(), weekStart: weekStartOf(date), ids: [], total: 0, byDate: {}, shiftByDate: {}, held: false, heldBy: '' };
      }
      const g = groups[key];
      const hm = /Supervisor OK by ([^<·]+?)\s*(?:<[^>]*>)?\s+on [^·]*awaiting senior approval/i.exec(String(f.Notes || ''));
      if (HELD_RE.test(String(f.Notes || ''))) { g.held = true; if (hm && !g.heldBy) g.heldBy = hm[1].trim(); }
      g.ids.push(it.id);
      if (f.Site) g.siteSet.add(String(f.Site).trim());
      const hr = Number(f.Hours) || 0;
      g.total += hr;
      if (!g.weekStart && weekStartOf(date)) g.weekStart = weekStartOf(date);
      if (!g.byDate[date]) g.byDate[date] = [];
      const allow = (String(f.RowType || '').toLowerCase() === 'allowance' || f.Allowance)
        ? (String(f.Allowance || '').trim() || 'Allowance')
        : '';
      g.byDate[date].push({ id: it.id, wo: String(f.WorkOrder || ''), hr, allow, comment: String(f.Comment || ''), jobs: allow ? [] : jobAllowancesOf(f) });
      const sh = (g.shiftByDate[date] = g.shiftByDate[date] || { worked: '', rostered: '' });   // staff app: same on every row of the day
      if (!sh.worked && f.ShiftWorked) sh.worked = String(f.ShiftWorked).trim();
      if (!sh.rostered && f.RosteredShift) sh.rostered = String(f.RosteredShift).trim();
    }

    const out = Object.values(groups).map(g => {
      const total = Math.round(g.total * 100) / 100;
      const contract = contracts[g.name.toLowerCase()] || 0;
      const siteList = [...g.siteSet].sort((a, b) => a.localeCompare(b));
      return {
      key: g.key,
      name: g.name,
      site: siteList.join(', '),
      weekStart: g.weekStart || '',
      total,
      contract,
      overBy: contract ? Math.max(0, Math.round((total - contract) * 100) / 100) : 0,
      held: g.held,
      heldBy: g.heldBy,
      canSenior: isSeniorFor(user, siteList, detailed),
      ids: g.ids,
      days: Object.keys(g.byDate).sort().map(date => ({ date, lines: g.byDate[date], shiftWorked: (g.shiftByDate[date] || {}).worked || '', rosteredShift: (g.shiftByDate[date] || {}).rostered || '' })),
    }; }).sort((a, b) =>
      (a.name.localeCompare(b.name)) || String(a.weekStart).localeCompare(String(b.weekStart))
    );

    return { statusCode: 200, headers, body: JSON.stringify({ ok: true, name: user.name, sites, groups: out, nextSend }) };
  } catch (err) {
    console.error(err);
    return { statusCode: 502, headers, body: '{"ok":false,"error":"server"}' };
  }
};
