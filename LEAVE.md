# Leave requests (design point — switched off until you turn it on)

Staff and contractors request leave in their timesheet app (a **Leave** tab appears when
it's on). Supervisors approve it in the Approvals app. Approved staff leave goes on the
Teams Shifts roster as **"Leave"** and shows on the staff timesheet for those days.

The same engine (`netlify/functions/leave-core.js`) is in all three apps — keep the copies
identical. `leave-ui.js` is shared by the staff and contractor apps.

## The switch

Either of these (the environment variable wins if set):

- **SharePoint** — in the **Settings** list add a row: Title `LeaveEnabled`, Value:
  - `Yes` → on for everyone · `No` → off
  - one or more emails (e.g. `a.mclennan@ael.co`) → on **only for those people** (testing).
    Contractors don't see it in this mode.
- **Netlify** — env `LEAVE_ENABLED` = `true` / `false` on each site (needs a redeploy).

Changes in SharePoint take effect within a minute.

## SharePoint lists to create (site AEL_Timesheets)

Type the column names exactly as shown (no spaces) when creating them.

**LeaveTypes** — one row per leave type (add more any time)

| Column | Type | Meaning |
|---|---|---|
| Title | (built in) | e.g. Annual Leave |
| Active | Yes/No | untick to retire a type |
| Site | Lookup → Sites, **allow multiple** | blank = every site |
| Entitlement | Yes/No | sick, carer's, compassionate, FDV… — no rules, supervisor **acknowledges** |
| Confidential | Yes/No | shown to others only as "Leave" (e.g. Family & Domestic Violence) |
| AllowPastDates | Yes/No | can be entered after the fact (sick) |
| NeedsAttachment | Yes/No | asks for a medical certificate / stat dec |
| SeniorApproval | Yes/No | always needs the senior approver (e.g. Long Service) |
| MinNoticeDays | Number | less notice → senior approver (e.g. Annual = 28) |
| MaxDaysBeforeSenior | Number | longer → senior approver (e.g. Annual = 10) |
| PayrollCode | Single line | for the future Xero link |
| Contractors | Yes/No | offered to contractors (none ticked → contractors get "Unavailable") |

Suggested rows: Annual Leave · Personal/Sick (Entitlement, AllowPastDates, NeedsAttachment) ·
Carer's (Entitlement, AllowPastDates) · Compassionate (Entitlement, AllowPastDates) ·
Family & Domestic Violence (Entitlement, Confidential, AllowPastDates) · Community Service
(Entitlement) · Long Service (SeniorApproval) · Parental (unpaid) · Unpaid · Time in Lieu.

**LeaveRequests** — written by the apps

| Column | Type |
|---|---|
| Title | (built in — the person's name) |
| Email, PersonType, PersonId, Site, LeaveType, StartDate, EndDate, Status, EvidenceUrl, ShiftsTimeOffId, Notified | Single line of text |
| Hours | Number (2 decimals) |
| DayHours, Note, RulesBroken, History | Multiple lines of text (plain) |

**Sites** — add **MaxOnLeave** (Number): the most people off at that site at once
(blank = no limit).

**Contractors** — optional **Email** column, so contractors are emailed decisions.

## How approval works

| Request | Supervisor taps Approve | Then |
|---|---|---|
| Contractor | Approved | — |
| Entitlement type | Acknowledged | — |
| Staff, no rule broken | Approved | — |
| Staff, a rule broken (senior-only type, short notice, too long, over MaxOnLeave) | "Supervisor OK" | the site's **SeniorApprover** approves → Approved |

A senior approver's own Approve is final. Nobody can approve their own leave. "Too many
off" is re-checked when the supervisor approves. Declining needs a reason; the person is
emailed either way (SENDER_EMAIL on the supervisor site).

Approvers are emailed about waiting leave every 10 minutes (`leave-notify`, supervisor
app). Confidential types are just called "leave" in emails and on the roster.

Leave days skip weekends and the site's public holidays; hours = Workers.ContractHours ÷ 5
(8 if not set). One-day requests can be part days.

## Teams Shifts

On approval, staff leave is added to the person's Shifts schedule (site **ShiftsTeamId**) as
a time-off block with the reason **Leave** (created automatically). Needs Graph
**Schedule.ReadWrite.All** (already granted). Cancelling future leave removes it. If Shifts
refuses (e.g. the person isn't in the team) the leave is still approved and the reason is
written to the request's History.

## Evidence

Certificates are saved to the site's **Documents** library under `LeaveEvidence/<year>/`
— restrict that folder's permissions to HR/supervisors.

## Not yet (future)

Leave balances and sending approved leave to **Xero** (PayrollCode is ready for it); leave
on the payroll report and in the over-contract check; supervisor-booked leave (shutdowns).
