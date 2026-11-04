// Payroll admin (payroll.html): pick a site and pay week, see everyone's hours, allowances
// and leave, chase who's outstanding, correct approved lines, download the Excel / PDF, and
// mark the week processed. One endpoint; the body says { action, ... }.
//
// Who can use it: emails in the Settings row Title "PayrollAdmins" (Value, separated by
// ; , or spaces), env PAYROLL_ADMINS, or a senior approver (env SENIOR_APPROVERS).
//
// Processed = a marker added to each Approved row's Notes ("Processed for payroll by … on …").
// Rows approved after that have no marker, so they show as new since processed. A processed
// row can't be changed here until the week is reopened. Payroll's own changes are stamped on
// the row's Notes too ("Payroll changed …"); a removed line is noted on the person's other rows.
const S = require('./supervisor');
const L = require('./leave-core');
const { buildReportData, makeXlsx, makePdf } = require('./report');

const PROC_RE = /Processed for payroll by ([^·]+?) on (\d{4}-\d{2}-\d{2} \d{2}:\d{2})/;
const PROC_STRIP = /(\s*·\s*)?Processed for payroll by [^·]*/g;
const LOG_RE = /Payroll (?:changed|removed|added|reopened)[^·]*/gi;
const dayOf = f => String(f.EntryDate || '').slice(0, 10);
const isAllow = f => String(f.RowType || '').toLowerCase() === 'allowance' || !!f.Allowance;
const r2 = n => Math.round((Number(n) || 0) * 100) / 100;
const auth = token => ({ Authorization: 'Bearer ' + token });

// Every page of a Graph list query (SharePoint hands back at most 999 rows a time).
async function allItems(token, url, prefer) {
  const out = [];
  for (let next = url, n = 0; next && n < 60; n++) {
    const r = await fetch(next, { headers: { ...auth(token), ...(prefer ? { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' } : {}) } });
    if (!r.ok) throw new Error('graph ' + r.status + ' ' + (await r.text()).slice(0, 300));
    const j = await r.json();
    out.push(...(j.value || []));
    next = j['@odata.nextLink'];
  }
  return out;
}
// Lists read several times in one request (the overview reads them per site) are kept for 20 s.
const _memo = new Map();
function memo(key, fn) {
  const hit = _memo.get(key);
  if (hit && Date.now() - hit.at < 20000) return hit.p;
  const p = fn().catch(e => { _memo.delete(key); throw e; });
  _memo.set(key, { at: Date.now(), p });
  return p;
}
function listByName(token, name) {
  return memo('list:' + name, async () => {
    const id = await S.findListId(token, name).catch(() => null);
    return id ? allItems(token, `${S.graphBase()}/lists/${id}/items?expand=fields&$top=999`).catch(() => []) : [];
  });
}

// ---- who may use it (Settings row re-read at most once a minute) ----
let _adm = null, _admAt = 0;
async function payrollAdmins(token) {
  if (_adm && Date.now() - _admAt < 60000) return _adm;
  const split = s => String(s || '').split(/[;,\s]+/).map(x => x.trim().toLowerCase()).filter(x => x.includes('@'));
  const set = new Set(split(process.env.PAYROLL_ADMINS));
  try {
    const row = (await listByName(token, 'Settings')).map(it => it.fields || {})
      .find(f => String(f.Title || '').trim().toLowerCase() === 'payrolladmins');
    if (row) for (const e of split(row.Value)) set.add(e);
  } catch (e) { console.warn('payroll admins setting', e.message); }
  _adm = set; _admAt = Date.now();
  return set;
}
async function isAdmin(token, email) {
  const e = String(email || '').trim().toLowerCase();
  return !!e && (S.isSenior(e) || (await payrollAdmins(token)).has(e));
}

// Active sites: [{ name, id, closeDayIndex }]
async function getSites(token) {
  const WEEKDAYS = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
  return (await listByName(token, 'Sites')).map(it => {
    const f = it.fields || {};
    const idx = WEEKDAYS.indexOf(String(f.CloseDay || '').trim().toLowerCase());
    return { name: String(f.Title || '').trim(), id: String(it.id), closeDayIndex: idx < 0 ? 0 : idx, active: f.Active !== false };
  }).filter(s => s.name && s.active).sort((a, b) => a.name.localeCompare(b.name));
}

// Timesheets rows dated lo..hi (ISO dates). Filters on EntryDate when SharePoint allows it,
// otherwise reads everything and filters here.
async function rowsBetween(token, lo, hi) {
  const from = L.isoOf(L.msOf(lo) - 86400000), to = L.isoOf(L.msOf(hi) + 86400000);
  let rows;
  try {
    rows = await allItems(token, `${S.itemsUrl()}?expand=fields&$top=999&$filter=fields/EntryDate ge '${from}T00:00:00Z' and fields/EntryDate le '${to}T23:59:59Z'`, true);
  } catch (e) {
    console.warn('payroll: date filter refused, reading all rows —', e.message);
    rows = await allItems(token, `${S.itemsUrl()}?expand=fields&$top=999`);
  }
  return rows.filter(it => { const d = dayOf(it.fields || {}); return d >= lo && d <= hi; });
}
const forSiteWeek = (rows, site, ws, we) => rows.filter(it => {
  const f = it.fields || {}; const d = dayOf(f);
  return d >= ws && d <= we && String(f.Site || '').trim() === site;
});

// The people expected to put in a timesheet at a site: Workers (staff, by home site) and
// active Contractors. Site may be text/choice or a lookup on either list.
async function expectedPeople(token, site) {
  const [sites, workers, contractors] = await Promise.all([getSites(token), listByName(token, 'Workers'), listByName(token, 'Contractors')]);
  const siteId = (sites.find(s => s.name === site) || {}).id;
  const atSite = (f, textKeys, idKeys) => {
    for (const k of textKeys) { const v = f[k]; if (typeof v === 'string' && v.trim()) return v.split(/[;,]/).map(x => x.trim().toLowerCase()).includes(site.toLowerCase()); }
    for (const k of idKeys) { const v = f[k]; const ids = Array.isArray(v) ? v : (v != null && v !== '' ? [v] : []); if (ids.length) return ids.map(String).includes(String(siteId)); }
    return false;
  };
  const out = [];
  for (const it of workers) {
    const f = it.fields || {}; const name = String(f.Title || '').trim();
    if (!name || f.Active === false || !atSite(f, ['HomeSite', 'Site'], ['HomeSiteLookupId', 'SiteLookupId'])) continue;
    out.push({ name, type: 'Staff', id: '', email: String(f.Email || '').trim() || S.emailFromName(name) });
  }
  for (const it of contractors) {
    const f = it.fields || {}; const name = String(f.Title || f.Name || '').trim();
    if (!name || f.Active === false || !atSite(f, ['Site'], ['SiteLookupId'])) continue;
    out.push({ name, type: 'Contractor', id: String(it.id), email: String(f.Email || '').trim() });
  }
  return out;
}

// Leave that touches the week at this site: [{ name, email, type, code, status, days:[{date,hours}] }]
async function leaveFor(token, site, ws, we) {
  try {
    const [{ requests }, types] = await Promise.all([memo('leave', () => L.getRequests(token)), memo('leavetypes', () => L.getLeaveTypes(token))]);
    const codeOf = t => ((types.find(x => x.title === t) || {}).code || '');
    return requests.filter(r => r.site === site && ['Pending', 'Supervisor OK', ...L.FINAL_OK].includes(r.status))
      .map(r => {
        const all = L.daysOf(r);
        const days = all.filter(d => d >= ws && d <= we).map(d => ({
          date: d, hours: r2(r.dayHours[d] != null ? r.dayHours[d] : (r.hours && all.length ? r.hours / all.length : 0)),
        }));
        return { name: r.name, email: r.email, type: r.type, code: codeOf(r.type), status: r.status, days };
      }).filter(r => r.days.length);
  } catch (e) { console.warn('payroll leave', e.message); return []; }
}

// ---- one site's week, shaped per person ----
// ---- employee lookup: everyone who can have a timesheet, and one person's history ----
// Sites list, every row (inactive too) → { id: name }
async function siteNamesById(token) {
  const out = {};
  for (const it of await listByName(token, 'Sites')) { const n = String((it.fields || {}).Title || '').trim(); if (n) out[String(it.id)] = n; }
  return out;
}
const lookupNames = (v, byId) => (Array.isArray(v) ? v : (v != null && v !== '' ? [v] : [])).map(x => byId[String(x && typeof x === 'object' ? x.LookupId : x)]).filter(Boolean);
async function allPeople(token) {
  const [byId, workers, contractors] = await Promise.all([siteNamesById(token), listByName(token, 'Workers'), listByName(token, 'Contractors')]);
  const siteOf = (f, textKeys, idKeys) => {
    for (const k of textKeys) if (typeof f[k] === 'string' && f[k].trim()) return f[k].trim();
    for (const k of idKeys) { const n = lookupNames(f[k], byId); if (n.length) return n.join(', '); }
    return '';
  };
  const out = [];
  for (const it of workers) {
    const f = it.fields || {}; const name = String(f.Title || '').trim(); if (!name) continue;
    const ch = parseFloat(f.ContractHours);
    out.push({ name, type: 'Staff', id: '', email: String(f.Email || '').trim() || S.emailFromName(name),
      site: siteOf(f, ['HomeSite', 'Site'], ['HomeSiteLookupId', 'SiteLookupId']), active: f.Active !== false, contractHours: ch > 0 ? ch : 0 });
  }
  for (const it of contractors) {
    const f = it.fields || {}; const name = String(f.Title || f.Name || '').trim(); if (!name) continue;
    out.push({ name, type: 'Contractor', id: String(it.id), email: String(f.Email || '').trim(),
      site: siteOf(f, ['Site'], ['SiteLookupId']), active: f.Active !== false, contractHours: 0 });
  }
  return out.sort((a, b) => a.name.localeCompare(b.name) || a.type.localeCompare(b.type));
}
// Is this Timesheets row the person's? By ContractorId (contractors), name, or the company email.
function rowIsPerson(f, who) {
  if (who.id && String(f.ContractorId || '').trim() === who.id) return true;
  const n = String(f.Title || '').trim().toLowerCase();
  if (n && n === who.name.toLowerCase()) return true;
  const e = S.emailFromName(f.Title);
  return !!(e && who.email && e === who.email.toLowerCase());
}
// One person's rows lo..hi, grouped by pay week, with everything the row holds.
function personWeeks(rows, leave) {
  const weeks = {};
  for (const it of rows) {
    const f = it.fields || {};
    const ws = S.payWeekOf(f) || dayOf(f);
    const w = weeks[ws] || (weeks[ws] = { weekStart: ws, weekEnd: L.isoOf(L.msOf(ws) + 6 * 86400000), lines: [], sites: new Set(), leave: [] });
    const notes = String(f.Notes || '');
    const pm = PROC_RE.exec(notes);
    if (f.Site) w.sites.add(String(f.Site).trim());
    w.lines.push({
      id: String(it.id), date: dayOf(f), site: String(f.Site || '').trim(), status: String(f.Status || 'Submitted').trim(),
      allow: isAllow(f) ? (String(f.Allowance || f.WorkOrder || '').trim() || 'Allowance') : '',
      wo: isAllow(f) ? '' : String(f.WorkOrder || '').trim(), hr: isAllow(f) ? 0 : r2(f.Hours),
      jobs: isAllow(f) ? [] : S.jobAllowancesOf(f), comment: String(f.Comment || ''),
      shiftWorked: String(f.ShiftWorked || '').trim(), rostered: String(f.RosteredShift || '').trim(),
      approvedBy: String(f.ApprovedBy || '').replace(/\s*<[^>]*>/, '').trim(), approvedDate: String(f.ApprovedDate || ''),
      rejection: String(f.RejectionReason || '').trim(), held: S.HELD_RE.test(notes),
      processed: pm ? { by: pm[1].trim(), when: pm[2] } : null,
      notes: notes.split(/\s+·\s+/).map(x => x.trim()).filter(Boolean),
      created: String(it.createdDateTime || ''), modified: String(it.lastModifiedDateTime || ''),
      batch: String(f.BatchID || '').trim(),
    });
  }
  for (const l of leave) {
    for (const d of l.days) {
      const w = Object.values(weeks).find(x => d.date >= x.weekStart && d.date <= x.weekEnd);
      if (w) { let e = w.leave.find(x => x.type === l.type && x.status === l.status); if (!e) w.leave.push(e = { type: l.type, code: l.code, status: l.status, days: [] }); e.days.push(d); }
    }
  }
  return Object.values(weeks).map(w => {
    w.lines.sort((a, b) => a.date.localeCompare(b.date) || (a.allow ? 1 : 0) - (b.allow ? 1 : 0));
    const has = st => w.lines.some(l => l.status === st);
    const ap = w.lines.filter(l => l.status === 'Approved');
    return {
      ...w, sites: [...w.sites].sort(),
      status: has('Rejected') && !has('Submitted') ? 'sentback' : w.lines.some(l => l.status === 'Submitted' && l.held) ? 'senior' : has('Submitted') ? 'waiting' : 'approved',
      total: r2(w.lines.filter(l => l.status !== 'Rejected').reduce((t, l) => t + l.hr, 0)),
      processed: { total: ap.length, done: ap.filter(l => l.processed).length },
    };
  }).sort((a, b) => b.weekStart.localeCompare(a.weekStart));
}

// ---- allowances: what kind each one is, and a person × allowance count for the week ----
// Allowances list → { title: 'daily' | 'weekly' | 'job' } (Frequency: Daily (default), Weekly, Per job)
function allowanceTypes(token) {
  return memo('allowtypes', async () => {
    const out = {};
    for (const it of await listByName(token, 'Allowances')) {
      const f = it.fields || {}; if (f.Active === false) continue;
      const t = String(f.Title || '').trim(); if (!t || out[t]) continue;
      const fr = String(f.Frequency || 'Daily').trim().toLowerCase();
      out[t] = /job/.test(fr) ? 'job' : fr === 'weekly' ? 'weekly' : 'daily';
    }
    return out;
  });
}
const KIND_ORDER = { daily: 0, job: 1, weekly: 2 };
// Approved rows → { cols:[{name,kind}], people:[{name, counts:{}, days:{}, shifts:{}}], totals:{} }.
// Daily/weekly allowances are their own rows; per-job ones (Call-out) are ticks on an hours row;
// the shift worked is counted once per day.
function allowanceSummary(rows, types) {
  const by = {}, jobNames = new Set();
  for (const it of rows) {
    const f = it.fields || {};
    const name = String(f.Title || '').trim() || 'Worker';
    const p = by[name.toLowerCase()] || (by[name.toLowerCase()] = { name, counts: {}, days: {}, shifts: {}, shiftDays: new Set() });
    const d = dayOf(f);
    const bump = k => { p.counts[k] = (p.counts[k] || 0) + 1; (p.days[k] = p.days[k] || []).push(d); };
    if (isAllow(f)) bump(String(f.Allowance || f.WorkOrder || '').trim() || 'Allowance');
    else {
      for (const j of S.jobAllowancesOf(f)) { jobNames.add(j); bump(j); }
      const sh = String(f.ShiftWorked || '').trim();
      if (sh && !p.shiftDays.has(d)) { p.shiftDays.add(d); p.shifts[sh] = (p.shifts[sh] || 0) + 1; }
    }
  }
  const people = Object.values(by).map(({ shiftDays, ...p }) => p).sort((a, b) => a.name.localeCompare(b.name));
  const names = [...new Set(people.flatMap(p => Object.keys(p.counts)))];
  const cols = names.map(n => ({ name: n, kind: types[n] || (jobNames.has(n) ? 'job' : 'daily') }))
    .sort((a, b) => KIND_ORDER[a.kind] - KIND_ORDER[b.kind] || a.name.localeCompare(b.name));
  const totals = {};
  for (const p of people) for (const [k, n] of Object.entries(p.counts)) totals[k] = (totals[k] || 0) + n;
  return { cols, people: people.filter(p => Object.keys(p.counts).length || Object.keys(p.shifts).length), totals };
}

function personKeyOf(f) { return String(f.ContractorId || f.Title || '').trim() || 'unknown'; }
function shapeWeek(site, ws, we, rows, expected, leave) {
  const people = {};
  const nameKey = n => String(n || '').trim().toLowerCase();
  const byName = {}, byEmail = {};
  // one person, however they're spelt: the same name, or the same company email
  // (staff rows carry no email, so theirs comes from the name: first-initial.surname@ael.co)
  const find = (name, email) => byName[nameKey(name)] || byEmail[String(email || S.emailFromName(name)).toLowerCase()] || null;
  const add = (key, name, email) => {
    if (!people[key]) {
      people[key] = { key, name, type: '', email: '', lines: [], leave: [], log: [] };
      byName[nameKey(name)] = people[key];
      const e = String(email || S.emailFromName(name)).toLowerCase();
      if (e && !byEmail[e]) byEmail[e] = people[key];
    }
    return people[key];
  };
  for (const it of rows) {
    const f = it.fields || {};
    const p = find(f.Title) || add(personKeyOf(f), String(f.Title || '').trim() || 'Worker');
    const notes = String(f.Notes || '');
    const pm = PROC_RE.exec(notes);
    p.lines.push({
      id: String(it.id), date: dayOf(f), status: String(f.Status || 'Submitted').trim(),
      allow: isAllow(f) ? (String(f.Allowance || f.WorkOrder || '').trim() || 'Allowance') : '',
      wo: isAllow(f) ? '' : String(f.WorkOrder || '').trim(), hr: isAllow(f) ? 0 : r2(f.Hours),
      comment: String(f.Comment || ''), jobs: isAllow(f) ? [] : S.jobAllowancesOf(f),
      shift: String(f.ShiftWorked || '').trim(), rostered: String(f.RosteredShift || '').trim(),
      held: S.HELD_RE.test(notes), processed: pm ? { by: pm[1].trim(), when: pm[2] } : null,
      approvedBy: String(f.ApprovedBy || '').trim(), rejection: String(f.RejectionReason || '').trim(),
    });
    for (const m of notes.match(LOG_RE) || []) if (!p.log.includes(m.trim())) p.log.push(m.trim());
  }
  for (const e of expected) {
    const p = find(e.name, e.email) || add(e.id || e.name, e.name, e.email);
    p.type = p.type || e.type; p.email = p.email || e.email;
  }
  for (const l of leave) {
    const p = find(l.name, l.email) || add(l.email || l.name, l.name, l.email);
    if (!p.email) p.email = l.email;
    p.leave.push({ type: l.type, code: l.code, status: l.status, days: l.days });
  }
  const list = Object.values(people).map(p => {
    const has = s => p.lines.some(l => l.status === s);
    const status = !p.lines.length ? 'missing'
      : has('Rejected') && !has('Submitted') ? 'sentback'
      : p.lines.some(l => l.status === 'Submitted' && l.held) ? 'senior'
      : has('Submitted') ? 'waiting' : 'approved';
    const ap = p.lines.filter(l => l.status === 'Approved');
    const processed = !ap.length ? 'no' : ap.every(l => l.processed) ? 'yes' : ap.some(l => l.processed) ? 'partial' : 'no';
    p.lines.sort((a, b) => a.date.localeCompare(b.date) || (a.allow ? 1 : 0) - (b.allow ? 1 : 0));
    return {
      ...p, status, processed,
      total: r2(p.lines.filter(l => l.status !== 'Rejected').reduce((s, l) => s + l.hr, 0)),
      approvedTotal: r2(ap.reduce((s, l) => s + l.hr, 0)),
      leaveHours: r2(p.leave.filter(l => L.FINAL_OK.includes(l.status)).reduce((s, l) => s + l.days.reduce((t, d) => t + d.hours, 0), 0)),
    };
  }).sort((a, b) => a.name.localeCompare(b.name));
  const ap = rows.filter(it => (it.fields || {}).Status === 'Approved');
  const stamped = ap.map(it => PROC_RE.exec(String((it.fields || {}).Notes || ''))).filter(Boolean);
  const last = stamped.sort((a, b) => b[2].localeCompare(a[2]))[0];
  const count = s => list.filter(p => p.status === s).length;
  return {
    site, weekStart: ws, weekEnd: we, people: list,
    summary: {
      people: list.length, approved: count('approved'), waiting: count('waiting') + count('senior'),
      lines: ap.length, processedLines: stamped.length,
      sentback: count('sentback'), missing: count('missing'),
      approvedHours: r2(list.reduce((s, p) => s + p.approvedTotal, 0)),
      leaveHours: r2(list.reduce((s, p) => s + p.leaveHours, 0)),
    },
    processed: {
      state: !ap.length ? 'no' : stamped.length === ap.length ? 'yes' : stamped.length ? 'partial' : 'no',
      by: last ? last[1].trim() : '', when: last ? last[2] : '', newSince: ap.length - stamped.length,
    },
  };
}

async function loadWeek(token, site, ws) {
  if (!L.validDate(ws)) throw Object.assign(new Error('Choose a week.'), { status: 400 });
  const we = L.isoOf(L.msOf(ws) + 6 * 86400000);
  const [rows, expected, leave] = await Promise.all([rowsBetween(token, ws, we), expectedPeople(token, site), leaveFor(token, site, ws, we)]);
  const mine = forSiteWeek(rows, site, ws, we);
  return { rows: mine, week: shapeWeek(site, ws, we, mine, expected, leave), leave };
}

async function createItem(token, fields) {
  const post = body => fetch(S.itemsUrl(), { method: 'POST', headers: { ...auth(token), 'Content-Type': 'application/json' }, body: JSON.stringify({ fields: body }) });
  let r = await post(fields);
  if (!r.ok && r.status === 400) {                       // optional columns missing → keep the essentials
    const { ApprovedBy, ApprovedDate, Comment, ...rest } = fields;
    r = await post(rest);
  }
  if (!r.ok) throw new Error('graph create ' + r.status + ' ' + (await r.text()).slice(0, 300));
  return r.json();
}
const appendNote = (prev, add) => { const p = String(prev || '').trim(); return p ? p + ' · ' + add : add; };

exports.handler = async (event) => {
  const headers = { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' };
  const reply = (code, obj) => ({ statusCode: code, headers, body: JSON.stringify(obj) });
  const user = await S.validateSupervisorToken(event.headers && (event.headers.authorization || event.headers.Authorization));
  if (!user) return reply(401, { ok: false });
  let data; try { data = JSON.parse(event.body || '{}'); } catch { return reply(400, { ok: false, error: 'bad json' }); }

  try {
    const action = String(data.action || 'me');
    if (action !== 'overview') _memo.clear();           // anything that writes reads fresh lists
    const token = await S.getAppToken();
    if (!(await isAdmin(token, user.email))) return reply(200, { ok: true, admin: false, name: user.name, email: user.email });
    const sites = await getSites(token);
    const site = String(data.site || '').trim();
    const ws = String(data.weekStart || '').trim();
    const siteOk = () => { if (!sites.some(s => s.name === site)) throw Object.assign(new Error('Choose a site.'), { status: 400 }); };
    const nowIso = new Date().toISOString();
    const when = nowIso.slice(0, 16).replace('T', ' ');
    const who = user.name || user.email;

    // ---- sign-in check + the sites, each with its recent pay weeks ----
    if (action === 'me') {
      return reply(200, {
        ok: true, admin: true, name: user.name, email: user.email,
        sites: sites.map(s => ({ name: s.name, weeks: Array.from({ length: 13 }, (_, n) => ({ ...S.weekForClose(s.closeDayIndex, n), current: n === 0 })) })),
      });
    }

    // ---- home page: every site's last completed week at a glance ----
    if (action === 'overview') {
      const back = Number(data.weeksBack) >= 0 ? Number(data.weeksBack) : 1;
      const weeks = sites.map(s => ({ s, ...S.weekForClose(s.closeDayIndex, back) }));
      if (!weeks.length) return reply(200, { ok: true, sites: [] });
      const lo = weeks.map(w => w.weekStart).sort()[0], hi = weeks.map(w => w.weekEnd).sort().pop();
      const rows = await rowsBetween(token, lo, hi);
      const out = await Promise.all(weeks.map(async w => {
        const [expected, leave] = await Promise.all([expectedPeople(token, w.s.name), leaveFor(token, w.s.name, w.weekStart, w.weekEnd)]);
        const wk = shapeWeek(w.s.name, w.weekStart, w.weekEnd, forSiteWeek(rows, w.s.name, w.weekStart, w.weekEnd), expected, leave);
        return { site: w.s.name, weekStart: w.weekStart, weekEnd: w.weekEnd, summary: wk.summary, processed: wk.processed };
      }));
      return reply(200, { ok: true, sites: out });
    }

    // ---- employee lookup: the list of people ----
    if (action === 'people') return reply(200, { ok: true, people: await allPeople(token) });

    // ---- employee lookup: one person's timesheets, past or present ----
    // { person:{ name, id?, email? }, from, to }  (ISO dates; at most a year at a time)
    if (action === 'person') {
      const q = data.person || {};
      const who = { name: String(q.name || '').trim(), id: String(q.id || '').trim(), email: String(q.email || '').trim() };
      if (!who.name) return reply(400, { ok: false, error: 'Choose a person.' });
      const to = L.validDate(data.to) ? data.to : L.isoOf(Date.now());
      let from = L.validDate(data.from) ? data.from : L.isoOf(L.msOf(to) - 83 * 86400000);
      if (L.msOf(to) - L.msOf(from) > 371 * 86400000) from = L.isoOf(L.msOf(to) - 371 * 86400000);
      if (from > to) return reply(400, { ok: false, error: 'The start date is after the end date.' });
      const [rows, people, lv] = await Promise.all([
        rowsBetween(token, L.isoOf(L.msOf(from) - 6 * 86400000), L.isoOf(L.msOf(to) + 6 * 86400000)),   // a whole pay week either side
        allPeople(token),
        memo('leave', () => L.getRequests(token)).then(r => r.requests).catch(() => []),
      ]);
      const info = people.find(p => p.name.toLowerCase() === who.name.toLowerCase() && (!who.id || p.id === who.id)) || { ...who, type: '', site: '' };
      if (!who.email) who.email = info.email || S.emailFromName(who.name);
      const mine = rows.filter(it => rowIsPerson(it.fields || {}, who));
      const types = await memo('leavetypes', () => L.getLeaveTypes(token)).catch(() => []);
      const leave = lv.filter(r => (r.name.toLowerCase() === who.name.toLowerCase() || (who.email && r.email === who.email.toLowerCase())))
        .map(r => {
          const all = L.daysOf(r);
          return { type: r.type, code: ((types.find(t => t.title === r.type) || {}).code || ''), status: r.status, start: r.start, end: r.end, note: r.note,
            days: all.map(d => ({ date: d, hours: r2(r.dayHours[d] != null ? r.dayHours[d] : (r.hours && all.length ? r.hours / all.length : 0)) })) };
        });
      const inRange = leave.filter(r => r.end >= from && r.start <= to);
      const weeks = personWeeks(mine, inRange).filter(w => w.weekEnd >= from && w.weekStart <= to);
      return reply(200, {
        ok: true, person: { ...info, name: info.name || who.name, email: who.email }, from, to, weeks,
        leave: inRange.filter(r => r.status !== 'Cancelled'),
        totals: { hours: r2(weeks.reduce((t, w) => t + w.total, 0)), weeks: weeks.length },
      });
    }

    siteOk();

    // ---- one site's week ----
    if (action === 'week') {
      const [{ week }, allowanceKinds] = await Promise.all([loadWeek(token, site, ws), allowanceTypes(token).catch(() => ({}))]);
      const allowances = Object.keys(allowanceKinds).sort((a, b) => a.localeCompare(b));
      return reply(200, { ok: true, ...week, allowances, allowanceKinds });
    }

    // ---- download Excel / PDF of the approved hours (with leave) ----
    if (action === 'export') {
      const { rows, week, leave } = await loadWeek(token, site, ws);
      const approved = rows.filter(it => (it.fields || {}).Status === 'Approved');
      const rep = buildReportData(site, week.weekStart, week.weekEnd, approved);
      rep.allowSummary = allowanceSummary(approved, await allowanceTypes(token).catch(() => ({})));
      rep.leave = leave.filter(l => L.FINAL_OK.includes(l.status)).flatMap(l => l.days.map(d => ({ name: l.name, date: d.date, type: l.type, code: l.code, hours: d.hours })))
        .sort((a, b) => a.name.localeCompare(b.name) || a.date.localeCompare(b.date));
      const pdf = String(data.format) === 'pdf';
      const buf = pdf ? await makePdf(rep) : await makeXlsx(rep);
      const file = `Payroll_${site}_${week.weekStart}_to_${week.weekEnd}`.replace(/[^A-Za-z0-9_.-]/g, '_') + (pdf ? '.pdf' : '.xlsx');
      return reply(200, { ok: true, file, mime: pdf ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', data: buf.toString('base64') });
    }

    // ---- mark processed / reopen: the whole week, one person (person = key), or some lines (ids) ----
    if (action === 'process' || action === 'reopen') {
      const { rows } = await loadWeek(token, site, ws);
      const ids = Array.isArray(data.ids) ? new Set(data.ids.map(String)) : null;
      const person = data.person != null ? String(data.person) : null;
      const ap = rows.filter(it => (it.fields || {}).Status === 'Approved'
        && (!ids || ids.has(String(it.id))) && (person == null || personKeyOf(it.fields || {}) === person));
      const todo = ap.filter(it => PROC_RE.test(String((it.fields || {}).Notes || '')) === (action === 'reopen'));
      const done = [];
      for (let i = 0; i < todo.length; i += 5) {          // a few at a time: quicker, and gentle on SharePoint
        await Promise.all(todo.slice(i, i + 5).map(async it => {
          const notes = String((it.fields || {}).Notes || '');
          const next = action === 'process'
            ? appendNote(notes, `Processed for payroll by ${who} on ${when}`)
            : appendNote(notes.replace(PROC_STRIP, '').replace(/^\s*·\s*/, '').trim(), `Payroll reopened by ${who} on ${when}`);
          await S.updateItem(token, it.id, { Notes: next });
          done.push(String(it.id));
        }));
      }
      return reply(200, { ok: true, updated: done.length, ids: done, by: who, when });
    }

    // ---- payroll corrections to approved lines ----
    if (action === 'edit') {
      const { rows } = await loadWeek(token, site, ws);
      const we = L.isoOf(L.msOf(ws) + 6 * 86400000);
      const byId = new Map(rows.map(it => [String(it.id), it]));
      const edits = (data.edits && typeof data.edits === 'object') ? data.edits : {};
      const removed = [...new Set((Array.isArray(data.removedIds) ? data.removedIds : []).map(String))];
      const added = Array.isArray(data.added) ? data.added.slice(0, 50) : [];
      const note = String(data.note || '').trim().slice(0, 300);
      const tail = ` (by ${who} on ${when}${note ? ' — ' + note : ''})`;
      const editable = id => {
        const it = byId.get(String(id)); const f = (it && it.fields) || {};
        if (!it) return 'That line is no longer in this week — reload.';
        if (f.Status !== 'Approved') return 'Only approved lines can be changed here.';
        if (PROC_RE.test(String(f.Notes || ''))) return 'This week is marked processed — reopen it first.';
        return '';
      };
      for (const id of [...Object.keys(edits), ...removed]) { const why = editable(id); if (why) return reply(400, { ok: false, error: why }); }
      // check every added line before writing anything
      const toAdd = [];
      for (const a of added) {
        const date = String(a.date || '');
        if (!L.validDate(date) || date < ws || date > we) return reply(400, { ok: false, error: 'Pick a day in this week.' });
        const tpl = rows.find(it => personKeyOf(it.fields || {}) === String(a.person || '') && (it.fields || {}).Status === 'Approved' && !PROC_RE.test(String((it.fields || {}).Notes || '')));
        if (!tpl) return reply(400, { ok: false, error: 'Lines can only be added for someone with approved, unprocessed hours this week.' });
        const t = tpl.fields || {};
        const base = {
          Title: t.Title, ContractorId: t.ContractorId, Site: site, BatchID: t.BatchID, EntryDate: date + 'T00:00:00Z',
          Status: 'Approved', ApprovedBy: who, ApprovedDate: nowIso,
        };
        let fields;
        if (a.kind === 'allowance') {
          const name = String(a.allow || '').trim().slice(0, 200);
          if (!name) return reply(400, { ok: false, error: 'Choose the allowance.' });
          fields = { ...base, RowType: 'Allowance', Allowance: name, Hours: 0, Notes: `Payroll added ${date} ${name}${tail}` };
        } else {
          const h = parseFloat(a.hr);
          if (isNaN(h) || h <= 0 || h > 24) return reply(400, { ok: false, error: 'Hours must be between 0 and 24.' });
          const wo = String(a.wo || '').trim().slice(0, 200);
          fields = { ...base, WorkOrder: wo, Hours: r2(h), Notes: `Payroll added ${date} ${wo || 'line'} ${r2(h)} h${tail}` };
        }
        Object.keys(fields).forEach(k => fields[k] == null && delete fields[k]);
        toAdd.push(fields);
      }

      let changed = 0, removedN = 0, addedN = 0;
      const removedNotes = {};                        // person key → what was removed
      for (const [id, e] of Object.entries(edits)) {
        if (removed.includes(id)) continue;
        const f = byId.get(id).fields || {};
        const upd = {}, what = [];
        if (e.hr != null && !isAllow(f)) {
          const h = parseFloat(e.hr);
          if (isNaN(h) || h < 0 || h > 24) return reply(400, { ok: false, error: 'Hours must be between 0 and 24.' });
          if (r2(h) !== r2(f.Hours)) { upd.Hours = r2(h); what.push(`hours ${r2(f.Hours)} → ${r2(h)}`); }
        }
        if (e.wo != null && !isAllow(f)) {
          const wo = String(e.wo).trim().slice(0, 200);
          if (wo !== String(f.WorkOrder || '').trim()) { upd.WorkOrder = wo; what.push(`work order “${String(f.WorkOrder || '').trim()}” → “${wo}”`); }
        }
        if (e.comment != null) {
          const c = String(e.comment).trim().slice(0, 255);
          if (c !== String(f.Comment || '').trim()) { upd.Comment = c; what.push('comment'); }
        }
        if (!what.length) continue;
        upd.Notes = appendNote(f.Notes, `Payroll changed ${dayOf(f)} ${what.join(', ')}${tail}`);
        try { await S.updateItem(token, id, upd); }
        catch (err) { const { Comment, ...rest } = upd; if (Comment == null) throw err; await S.updateItem(token, id, rest); }
        changed++;
      }
      for (const id of removed) {
        const f = byId.get(id).fields || {};
        await S.deleteItem(token, id); removedN++;
        const k = personKeyOf(f);
        (removedNotes[k] = removedNotes[k] || []).push(`${dayOf(f)} ${isAllow(f) ? (f.Allowance || 'allowance') : `${String(f.WorkOrder || '').trim() || 'line'} ${r2(f.Hours)} h`}`);
      }
      for (const fields of toAdd) { await createItem(token, fields); addedN++; }
      // note removals on one of the person's remaining approved rows, so there's a record
      for (const [k, items] of Object.entries(removedNotes)) {
        const keep = rows.find(it => personKeyOf(it.fields || {}) === k && (it.fields || {}).Status === 'Approved' && !removed.includes(String(it.id)));
        if (!keep) continue;
        let notes = String((keep.fields || {}).Notes || '');
        if (edits[String(keep.id)]) {                       // its Notes were just re-saved — read them back
          const r = await fetch(`${S.itemsUrl()}/${keep.id}?expand=fields`, { headers: auth(token) });
          if (r.ok) notes = String(((await r.json()).fields || {}).Notes || '');
        }
        await S.updateItem(token, keep.id, { Notes: appendNote(notes, `Payroll removed ${items.join('; ')}${tail}`) }).catch(e => console.warn('payroll removal note', e.message));
      }
      return reply(200, { ok: true, changed, removed: removedN, added: addedN });
    }

    return reply(400, { ok: false, error: 'unknown action' });
  } catch (err) {
    console.error(err);
    return reply(err.status || 502, { ok: false, error: err.status ? err.message : 'server' });
  }
};
