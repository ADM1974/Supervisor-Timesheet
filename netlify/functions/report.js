// Payroll reporting: for a site + pay week, build a spreadsheet (.xlsx) and a PDF
// of the APPROVED hours and email them to payroll (cc the site's supervisor).
//
// This file is BOTH the manual "send now" endpoint (a signed-in supervisor posts
// { which:'last'|'current', site? }) AND exports buildAndSendForSite() for the
// scheduled job. Email goes out via Microsoft Graph sendMail using the app-only
// token — needs the app to have Mail.Send (application) granted, and these env
// vars:  SENDER_EMAIL (mailbox to send from), PAYROLL_EMAIL (recipient).
const { getAppToken, validateSupervisorToken, getSupervisorSites, getSitesDetailed, getApprovedForSites, weekForClose } = require('./supervisor');
const ExcelJS = require('exceljs');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');

// ---- shape the approved rows into per-worker totals ----
function buildReportData(site, weekStart, weekEnd, rows) {
  const byWorker = {};
  for (const it of rows) {
    const f = it.fields || {};
    const key = String(f.ContractorId || f.Title || '').trim() || 'unknown';
    const w = byWorker[key] || (byWorker[key] = { name: String(f.Title || '').trim() || 'Worker', days: {}, total: 0, allowances: [], approvedBy: new Set() });
    const date = String(f.EntryDate || '').slice(0, 10);
    const isAllow = String(f.RowType || '').toLowerCase() === 'allowance' || !!f.Allowance;
    if (isAllow) {
      w.allowances.push({ date, name: String(f.Allowance || 'Allowance').trim() });
    } else {
      const hr = Number(f.Hours) || 0;
      (w.days[date] = w.days[date] || []).push({ wo: String(f.WorkOrder || '').trim(), hr });
      w.total += hr;
    }
    const ab = String(f.ApprovedBy || '').trim();
    if (ab) w.approvedBy.add(ab);
  }
  const workers = Object.values(byWorker).map(w => ({
    name: w.name,
    total: Math.round(w.total * 100) / 100,
    days: Object.keys(w.days).sort().map(date => ({ date, lines: w.days[date] })),
    allowances: w.allowances.sort((a, b) => a.date.localeCompare(b.date)),
    approvedBy: [...w.approvedBy].join(', '),
  })).sort((a, b) => a.name.localeCompare(b.name));
  const grandTotal = Math.round(workers.reduce((s, w) => s + w.total, 0) * 100) / 100;
  return { site, weekStart, weekEnd, workers, grandTotal };
}

// ---- spreadsheet ----
async function makeXlsx(rep) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(String(rep.site).slice(0, 28) || 'Payroll');
  ws.addRow([`Payroll — ${rep.site}`]).font = { bold: true, size: 14 };
  ws.addRow([`Pay week ${rep.weekStart} to ${rep.weekEnd} — APPROVED hours`]);
  ws.addRow([]);
  const head = ws.addRow(['Worker', 'Date', 'Work order', 'Hours', 'Allowance', 'Approved by']);
  head.font = { bold: true };
  for (const w of rep.workers) {
    for (const d of w.days) for (const l of d.lines) ws.addRow([w.name, d.date, l.wo, l.hr, '', w.approvedBy]);
    for (const a of w.allowances) ws.addRow([w.name, a.date, '', '', a.name, w.approvedBy]);
    ws.addRow([`${w.name} — total`, '', '', w.total, '', '']).font = { bold: true };
    ws.addRow([]);
  }
  ws.addRow(['GRAND TOTAL', '', '', rep.grandTotal, '', '']).font = { bold: true, size: 12 };
  ws.columns.forEach((c, i) => { c.width = i === 0 ? 24 : (i === 2 ? 26 : 16); });
  return Buffer.from(await wb.xlsx.writeBuffer());
}

// ---- PDF (positioned text via pdf-lib; standard fonts, no external files) ----
async function makePdf(rep) {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const bold = await doc.embedFont(StandardFonts.HelveticaBold);
  const M = 42; let page = doc.addPage([595, 842]); let y = 800;
  const put = (text, f = font, size = 10, indent = 0) => {
    if (y < 48) { page = doc.addPage([595, 842]); y = 800; }
    page.drawText(String(text), { x: M + indent, y, size, font: f, color: rgb(0.12, 0.12, 0.12) });
    y -= size + 6;
  };
  put(`Payroll — ${rep.site}`, bold, 16);
  put(`Pay week ${rep.weekStart} to ${rep.weekEnd} — approved hours`, font, 10);
  y -= 6;
  for (const w of rep.workers) {
    put(`${w.name}    (${w.total} h)`, bold, 12);
    for (const d of w.days) for (const l of d.lines) put(`${d.date}    ${l.wo || '-'}    ${l.hr} h`, font, 10, 14);
    for (const a of w.allowances) put(`${a.date}    ${a.name}    [allowance]`, font, 10, 14);
    if (w.approvedBy) put(`approved by ${w.approvedBy}`, font, 8, 14);
    y -= 6;
  }
  y -= 6;
  put(`GRAND TOTAL: ${rep.grandTotal} h`, bold, 13);
  return Buffer.from(await doc.save());
}

// ---- email via Graph sendMail (app-only; needs Mail.Send + SENDER_EMAIL/PAYROLL_EMAIL) ----
async function sendReport(token, rep, xlsxBuf, pdfBuf, ccEmails) {
  const sender = String(process.env.SENDER_EMAIL || '').trim();
  const payroll = String(process.env.PAYROLL_EMAIL || '').trim();
  if (!sender || !payroll) throw new Error('SENDER_EMAIL/PAYROLL_EMAIL not set');
  const base = `Payroll_${rep.site}_${rep.weekStart}_to_${rep.weekEnd}`.replace(/[^A-Za-z0-9_.-]/g, '_');
  const cc = [...new Set((ccEmails || []).map(e => String(e).trim().toLowerCase()).filter(Boolean).filter(e => e !== payroll.toLowerCase()))];
  const message = {
    subject: `Payroll — ${rep.site} — week ${rep.weekStart} to ${rep.weekEnd}`,
    body: { contentType: 'HTML', content: `<p>Approved hours for <b>${rep.site}</b>, pay week ${rep.weekStart} to ${rep.weekEnd}.</p><p>${rep.workers.length} worker(s) · <b>${rep.grandTotal} h</b> total. Spreadsheet and PDF attached.</p>` },
    toRecipients: [{ emailAddress: { address: payroll } }],
    ccRecipients: cc.map(a => ({ emailAddress: { address: a } })),
    attachments: [
      { '@odata.type': '#microsoft.graph.fileAttachment', name: base + '.xlsx', contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', contentBytes: xlsxBuf.toString('base64') },
      { '@odata.type': '#microsoft.graph.fileAttachment', name: base + '.pdf', contentType: 'application/pdf', contentBytes: pdfBuf.toString('base64') },
    ],
  };
  const r = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(sender)}/sendMail`, {
    method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ message, saveToSentItems: true }),
  });
  if (!r.ok) throw new Error('sendMail ' + r.status + ' ' + (await r.text()));
}

// ---- one site, one week: build files + send (used by manual endpoint AND scheduler) ----
async function buildAndSendForSite(token, site, weekStart, weekEnd, ccEmails) {
  const rows = await getApprovedForSites(token, [site], weekStart, weekEnd);
  const rep = buildReportData(site, weekStart, weekEnd, rows);
  if (!rep.workers.length) return { site, sent: false, reason: 'no approved hours', weekStart, weekEnd };
  const [xlsx, pdf] = [await makeXlsx(rep), await makePdf(rep)];
  await sendReport(token, rep, xlsx, pdf, ccEmails);
  return { site, sent: true, workers: rep.workers.length, total: rep.grandTotal, weekStart, weekEnd };
}

// ---- manual endpoint: a signed-in supervisor sends their site(s)' report ----
exports.handler = async (event) => {
  const headers = { 'Content-Type': 'application/json' };
  const user = await validateSupervisorToken(event.headers && (event.headers.authorization || event.headers.Authorization));
  if (!user) return { statusCode: 401, headers, body: '{"ok":false}' };
  let data; try { data = JSON.parse(event.body || '{}'); } catch { return { statusCode: 400, headers, body: '{"ok":false,"error":"bad json"}' }; }
  try {
    const token = await getAppToken();
    const sites = await getSupervisorSites(token, user.email);
    if (!sites.length) return { statusCode: 403, headers, body: '{"ok":false,"error":"no sites"}' };
    const detailed = await getSitesDetailed(token);
    const weeksBack = String(data.which || 'last').toLowerCase() === 'current' ? 0 : 1;
    const targets = data.site ? sites.filter(s => s === data.site) : sites;
    const results = [];
    for (const site of targets) {
      const info = detailed[site] || { closeDayIndex: 0, approverEmails: [] };
      const { weekStart, weekEnd } = weekForClose(info.closeDayIndex, weeksBack);
      try { results.push(await buildAndSendForSite(token, site, weekStart, weekEnd, info.approverEmails)); }
      catch (e) { results.push({ site, sent: false, error: String((e && e.message) || e) }); }
    }
    return { statusCode: 200, headers, body: JSON.stringify({ ok: true, results }) };
  } catch (err) {
    console.error(err);
    return { statusCode: 502, headers, body: JSON.stringify({ ok: false, error: String((err && err.message) || err) }) };
  }
};

exports.buildAndSendForSite = buildAndSendForSite;
exports.buildReportData = buildReportData;
exports.makeXlsx = makeXlsx;
exports.makePdf = makePdf;
