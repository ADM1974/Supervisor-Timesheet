// Leave requests — shared engine. The SAME file is in all three apps (staff, contractor,
// supervisor); keep the copies identical.
//
// SharePoint (site SP_SITE_ID):
//   LeaveTypes     Title, Active, Site (lookup, multi; blank = every site), Entitlement,
//                  Confidential, AllowPastDates, NeedsAttachment, SeniorApproval (yes/no),
//                  MinNoticeDays, MaxDaysBeforeSenior (number), PayrollCode (text),
//                  Contractors (yes/no — offered to contractors)
//   LeaveRequests  Title (person), Email, PersonType (Staff/Contractor), PersonId, Site,
//                  LeaveType, StartDate, EndDate (text yyyy-mm-dd), Hours (number),
//                  DayHours, Note, RulesBroken, History (multi-line text), Status,
//                  EvidenceUrl, ShiftsTimeOffId, Notified (text)
//   Sites          + MaxOnLeave (number) — most people off at once at the site
//   Settings       a row Title "LeaveEnabled", Value = Yes / No / emails (only those people)
//
// The on/off switch: env LEAVE_ENABLED (true/false) wins; otherwise the Settings row.
// Status flow: Pending → Approved | Declined, or Pending → Supervisor OK (a rule was
// broken, waiting for the senior approver) → Approved | Declined. Entitlement types
// (sick, carer's, compassionate, FDV…) can't be refused for operational reasons, so
// they go Pending → Acknowledged with no rules. Contractors: Pending → Approved, no rules.

const spBase = () => `https://graph.microsoft.com/v1.0/sites/${process.env.SP_SITE_ID}`;
const auth = token => ({ Authorization: 'Bearer ' + token });
const yes = v => v === true || /^(yes|true|1|on)$/i.test(String(v == null ? '' : v).trim());
const num = v => { const n = parseFloat(v); return isFinite(n) ? n : 0; };

const ACTIVE = ['Pending', 'Supervisor OK', 'Approved', 'Acknowledged'];   // count as "off"
const FINAL_OK = ['Approved', 'Acknowledged'];

// ---------- lists ----------
let _lists = null, _listsAt = 0;
async function findList(token, name) {
  const want = String(name).toLowerCase().replace(/\s+/g, '');
  const pick = () => _lists.find(l => String(l.displayName || '').toLowerCase().replace(/\s+/g, '') === want) || null;
  const load = async () => {
    const r = await fetch(`${spBase()}/lists?$select=id,displayName&$top=999`, { headers: auth(token) });
    if (!r.ok) throw new Error('lists ' + r.status);
    _lists = (await r.json()).value || []; _listsAt = Date.now();
  };
  if (!_lists || Date.now() - _listsAt > 600000) await load();
  // A list made since the last look (e.g. while setting leave up) — look again rather than wait 10 min.
  if (!pick() && Date.now() - _listsAt > 5000) await load();
  return pick();
}
async function listItems(token, listId) {
  const out = [];
  let url = `${spBase()}/lists/${listId}/items?expand=fields&$top=999`;
  while (url) {
    const r = await fetch(url, { headers: auth(token) });
    if (!r.ok) throw new Error('list items ' + r.status + ' ' + (await r.text()).slice(0, 200));
    const j = await r.json();
    out.push(...(j.value || []));
    url = j['@odata.nextLink'] || '';
  }
  return out;
}

// ---------- the switch ----------
let _sw = null, _swAt = 0;
async function leaveSwitch(token) {
  const env = String(process.env.LEAVE_ENABLED || '').trim();
  if (env) return { on: yes(env), testers: null };
  if (_sw && Date.now() - _swAt < 60000) return _sw;
  let sw = { on: false, testers: null };
  try {
    const list = await findList(token, 'Settings');
    if (list) {
      const row = (await listItems(token, list.id)).map(it => it.fields || {})
        .find(f => String(f.Title || '').trim().toLowerCase() === 'leaveenabled');
      if (row) {
        const v = [row.Value, row.Setting, row.Enabled, row.Text, row.Data].find(x => x != null && String(x).trim() !== '');
        const s = String(v == null ? '' : v).trim();
        if (s.includes('@')) sw = { on: true, testers: s.split(/[;,\s]+/).map(x => x.trim().toLowerCase()).filter(Boolean) };
        else sw = { on: yes(v), testers: null };
      }
    }
  } catch (e) { console.warn('leave switch', e.message); }
  _sw = sw; _swAt = Date.now();
  return sw;
}
// Is leave on for this person? (email may be '' for contractors: testers-only mode hides it from them)
function leaveOnFor(sw, email) {
  if (!sw || !sw.on) return false;
  if (!sw.testers) return true;
  return sw.testers.includes(String(email || '').trim().toLowerCase());
}

// ---------- sites ----------
// { siteName: { id, max, teamId } } from the Sites list (MaxOnLeave, ShiftsTeamId).
async function getLeaveSites(token) {
  const list = await findList(token, 'Sites');
  if (!list) return {};
  const out = {};
  for (const it of await listItems(token, list.id)) {
    const f = it.fields || {};
    const name = String(f.Title || '').trim();
    if (!name || f.Active === false) continue;
    out[name] = { id: String(it.id), max: num(f.MaxOnLeave), teamId: String(f.ShiftsTeamId || '').trim() };
  }
  return out;
}

// ---------- leave types ----------
function siteIdsOf(f) {
  const ids = [];
  const add = v => { if (v != null && v !== '') ids.push(String(v)); };
  if (Array.isArray(f.SiteLookupId)) f.SiteLookupId.forEach(add); else add(f.SiteLookupId);
  if (Array.isArray(f.Site)) f.Site.forEach(x => add(x && typeof x === 'object' ? x.LookupId : null));
  return [...new Set(ids)];
}
async function getLeaveTypes(token) {
  const list = await findList(token, 'LeaveTypes');
  if (!list) return [];
  return (await listItems(token, list.id)).map(it => it.fields || {})
    .filter(f => String(f.Title || '').trim() && f.Active !== false)
    .map(f => ({
      title: String(f.Title).trim(),
      siteIds: siteIdsOf(f),                       // [] = every site
      entitlement: yes(f.Entitlement), confidential: yes(f.Confidential),
      allowPast: yes(f.AllowPastDates), needsAttachment: yes(f.NeedsAttachment),
      senior: yes(f.SeniorApproval), minNotice: num(f.MinNoticeDays), maxDays: num(f.MaxDaysBeforeSenior),
      code: String(f.PayrollCode || '').trim(), contractors: yes(f.Contractors),
    }))
    .sort((a, b) => a.title.localeCompare(b.title));
}
function typesForSite(types, siteId) { return types.filter(t => !t.siteIds.length || t.siteIds.includes(String(siteId))); }
// What a type is called where other people can see it (roster, "who else is off").
const publicName = (type) => (type && type.confidential ? 'Leave' : (type ? type.title : 'Leave'));

// ---------- dates ----------
const pad = n => String(n).padStart(2, '0');
const isoOf = ms => { const d = new Date(ms); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`; };
const msOf = s => Date.parse(s + 'T00:00:00Z');
const validDate = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')) && !isNaN(msOf(s));
function tzForSite(site) {
  const s = String(site || '').toUpperCase();
  if (/BNE|OOL|CNS|TSV/.test(s)) return 'Australia/Brisbane';
  if (/PER/.test(s)) return 'Australia/Perth';
  if (/ADL/.test(s)) return 'Australia/Adelaide';
  if (/AKL|WLG|CHC|ZQN/.test(s)) return 'Pacific/Auckland';
  return 'Australia/Sydney';
}
function todayIn(tz) { return new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date()); }
// Every date from start to end (inclusive).
function datesBetween(start, end) {
  const out = [];
  for (let t = msOf(start); t <= msOf(end) && out.length < 400; t += 86400000) out.push(isoOf(t));
  return out;
}
// Working days (Mon–Fri) in the range, skipping public holidays ({date:name}).
function workingDays(start, end, holidays) {
  return datesBetween(start, end).filter(d => { const wd = new Date(msOf(d)).getUTCDay(); return wd >= 1 && wd <= 5 && !(holidays || {})[d]; });
}
// UTC instant of local midnight on dateStr in tz (DST-safe).
function zonedMidnightUtc(dateStr, tz) {
  const guess = msOf(dateStr);
  const off = t => {
    const p = Object.fromEntries(new Intl.DateTimeFormat('en-US', { timeZone: tz, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' })
      .formatToParts(new Date(t)).map(x => [x.type, x.value]));
    return Date.UTC(+p.year, +p.month - 1, +p.day, +p.hour, +p.minute, +p.second) - t;
  };
  let t = guess - off(guess);
  t = guess - off(t);
  return new Date(t).toISOString();
}

// ---------- requests ----------
function reqOf(it) {
  const f = it.fields || {};
  let dayHours = {};
  try { dayHours = JSON.parse(String(f.DayHours || '{}')) || {}; } catch (_) { dayHours = {}; }
  return {
    id: String(it.id), name: String(f.Title || '').trim(), email: String(f.Email || '').trim().toLowerCase(),
    personType: String(f.PersonType || 'Staff').trim(), personId: String(f.PersonId || '').trim(),
    site: String(f.Site || '').trim(), type: String(f.LeaveType || '').trim(),
    start: String(f.StartDate || '').slice(0, 10), end: String(f.EndDate || '').slice(0, 10),
    hours: num(f.Hours), dayHours, note: String(f.Note || ''), status: String(f.Status || 'Pending').trim(),
    rules: String(f.RulesBroken || '').split('\n').map(s => s.trim()).filter(Boolean),
    history: String(f.History || ''), evidenceUrl: String(f.EvidenceUrl || '').trim(),
    timeOffId: String(f.ShiftsTimeOffId || '').trim(), notified: String(f.Notified || '').trim(),
    created: it.createdDateTime || '',
  };
}
async function getRequests(token) {
  const list = await findList(token, 'LeaveRequests');
  if (!list) return { listId: null, requests: [] };
  return { listId: list.id, requests: (await listItems(token, list.id)).map(reqOf) };
}
// The days a request takes someone off (its saved working days, or Mon–Fri of its range).
function daysOf(r) { const k = Object.keys(r.dayHours || {}); return k.length ? k.sort() : workingDays(r.start, r.end, {}); }

async function createRequest(token, listId, fields) {
  const r = await fetch(`${spBase()}/lists/${listId}/items`, {
    method: 'POST', headers: { ...auth(token), 'Content-Type': 'application/json' }, body: JSON.stringify({ fields }),
  });
  if (!r.ok) throw new Error('create leave ' + r.status + ' ' + (await r.text()).slice(0, 300));
  return reqOf(await r.json());
}
async function updateRequest(token, listId, id, fields) {
  const r = await fetch(`${spBase()}/lists/${listId}/items/${id}/fields`, {
    method: 'PATCH', headers: { ...auth(token), 'Content-Type': 'application/json' }, body: JSON.stringify(fields),
  });
  if (!r.ok) throw new Error('update leave ' + r.status + ' ' + (await r.text()).slice(0, 300));
}
const stamp = (who, what) => `${new Date().toISOString().slice(0, 16).replace('T', ' ')} UTC — ${what}${who ? ' (' + who + ')' : ''}`;
const addHistory = (r, line) => (r.history ? r.history + '\n' : '') + line;

// ---------- rules (staff, non-entitlement types) ----------
// For each day, how many OTHER people at the site are already off (pending or approved).
function othersOff(requests, site, days, excludeId, excludeEmail) {
  const out = {};
  for (const d of days) out[d] = [];
  for (const r of requests) {
    if (r.id === excludeId || r.site !== site || !ACTIVE.includes(r.status)) continue;
    if (excludeEmail && r.email && r.email === excludeEmail) continue;
    for (const d of daysOf(r)) if (out[d]) out[d].push(r);
  }
  return out;
}
const fmtDay = d => new Date(msOf(d)).toLocaleDateString('en-AU', { weekday: 'short', day: 'numeric', month: 'short', timeZone: 'UTC' });
function evaluateRules({ type, start, days, today, site, siteMax, others }) {
  if (!type || type.entitlement) return [];
  const rules = [];
  if (type.senior) rules.push(`${type.title} always needs the senior approver`);
  if (type.minNotice > 0 && today) {
    const notice = Math.round((msOf(start) - msOf(today)) / 86400000);
    if (notice < type.minNotice) rules.push(`Only ${Math.max(notice, 0)} days' notice — ${type.title} needs ${type.minNotice}`);
  }
  if (type.maxDays > 0 && days.length > type.maxDays) rules.push(`${days.length} working days — more than ${type.maxDays} needs the senior approver`);
  if (siteMax > 0 && others) {
    const over = days.filter(d => (others[d] || []).length + 1 > siteMax);
    if (over.length) rules.push(`More than ${siteMax} off at ${site} on ${over.slice(0, 4).map(fmtDay).join(', ')}${over.length > 4 ? ` (+${over.length - 4} more)` : ''}`);
  }
  return rules;
}

// ---------- Teams Shifts ----------
async function shiftsFetch(token, url, opts) {
  const actAs = String(process.env.SHIFTS_ACT_AS || '').trim();
  const tries = actAs ? [{}, { 'MS-APP-ACTS-AS': actAs }] : [{}];
  let r;
  for (const extra of tries) {
    r = await fetch(url, { ...(opts || {}), headers: { ...auth(token), 'Content-Type': 'application/json', ...extra } });
    if (r.ok || (r.status !== 403 && r.status !== 400)) return r;
  }
  return r;
}
async function teamTz(token, teamId, fallback) {
  try {
    const r = await shiftsFetch(token, `https://graph.microsoft.com/v1.0/teams/${teamId}/schedule`);
    if (r.ok) { const tz = (await r.json()).timeZone; if (tz) { todayIn(tz); return tz; } }
  } catch (_) { /* fall back */ }
  return fallback;
}
async function leaveReasonId(token, teamId) {
  const base = `https://graph.microsoft.com/v1.0/teams/${teamId}/schedule/timeOffReasons`;
  const r = await shiftsFetch(token, base);
  if (!r.ok) throw new Error('time-off reasons ' + r.status + ' ' + (await r.text()).slice(0, 200));
  const found = ((await r.json()).value || []).find(x => x.isActive !== false && String(x.displayName || '').trim().toLowerCase() === 'leave');
  if (found) return found.id;
  const c = await shiftsFetch(token, base, { method: 'POST', body: JSON.stringify({ displayName: 'Leave', iconType: 'plane', isActive: true }) });
  if (!c.ok) throw new Error('add time-off reason ' + c.status + ' ' + (await c.text()).slice(0, 200));
  return (await c.json()).id;
}
// Approved leave → a "Leave" time-off block on the person's Shifts schedule (the type
// isn't shown: rosters are seen by the whole team). Returns the time-off id.
async function addTimeOff(token, teamId, userId, start, end, site) {
  const tz = await teamTz(token, teamId, tzForSite(site));
  const reason = await leaveReasonId(token, teamId);
  const body = { timeOffReasonId: reason, startDateTime: zonedMidnightUtc(start, tz), endDateTime: zonedMidnightUtc(isoOf(msOf(end) + 86400000), tz), theme: 'gray' };
  const r = await shiftsFetch(token, `https://graph.microsoft.com/v1.0/teams/${teamId}/schedule/timesOff`, {
    method: 'POST', body: JSON.stringify({ userId, sharedTimeOff: body, draftTimeOff: body }),
  });
  if (!r.ok) throw new Error('add time off ' + r.status + ' ' + (await r.text()).slice(0, 300));
  return (await r.json()).id;
}
async function removeTimeOff(token, teamId, timeOffId) {
  if (!teamId || !timeOffId) return;
  const r = await shiftsFetch(token, `https://graph.microsoft.com/v1.0/teams/${teamId}/schedule/timesOff/${timeOffId}`, { method: 'DELETE' });
  if (!r.ok && r.status !== 404) throw new Error('remove time off ' + r.status + ' ' + (await r.text()).slice(0, 200));
}

// ---------- evidence (medical certificate / stat dec) ----------
// Saved in the site's Documents library under LeaveEvidence/<year>/. data = base64.
async function uploadEvidence(token, who, fileName, mime, data) {
  const buf = Buffer.from(String(data || ''), 'base64');
  if (!buf.length) throw new Error('empty file');
  if (buf.length > 4 * 1024 * 1024) throw new Error('file too big (4 MB max)');
  const safe = s => String(s || '').replace(/[^A-Za-z0-9._-]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 60) || 'file';
  const path = `LeaveEvidence/${new Date().getUTCFullYear()}/${Date.now()}-${safe(who)}-${safe(fileName)}`;
  const r = await fetch(`${spBase()}/drive/root:/${path}:/content`, {
    method: 'PUT', headers: { ...auth(token), 'Content-Type': mime || 'application/octet-stream' }, body: buf,
  });
  if (!r.ok) throw new Error('upload ' + r.status + ' ' + (await r.text()).slice(0, 200));
  return (await r.json()).webUrl || '';
}

module.exports = {
  ACTIVE, FINAL_OK, yes, findList, listItems, leaveSwitch, leaveOnFor, getLeaveSites, getLeaveTypes, typesForSite,
  publicName, isoOf, msOf, validDate, tzForSite, todayIn, datesBetween, workingDays, zonedMidnightUtc,
  getRequests, daysOf, createRequest, updateRequest, stamp, addHistory, othersOff, evaluateRules, fmtDay,
  addTimeOff, removeTimeOff, uploadEvidence,
};
