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
| NeedsAttachment | Yes/No | a certificate / stat dec may be needed — when, is set by the next two columns (supervisors can also ask) |
| CertificateAfterDays | Number (optional) | needed when off **more than** this many working days. Blank = 1; 0 = every time |
| CertificateNextToDayOff | Yes/No (optional) | also needed for a single day next to a weekend or public holiday. Blank = Yes |
| SeniorApproval | Yes/No | always needs the senior approver (e.g. Long Service) |
| MinNoticeDays | Number | less notice → senior approver (e.g. Annual = 28) |
| MaxDaysBeforeSenior | Number | longer → senior approver (e.g. Annual = 10) |
| PayrollCode | Single line | for the future Xero link |
| Contractors | Yes/No | offered to contractors (none ticked → contractors get "Unavailable") |

Suggested rows (the AEL baseline — `LeaveTypes_baseline.xlsx`):

| Type | Settings | Why |
|---|---|---|
| Annual Leave | MinNoticeDays 28, MaxDaysBeforeSenior 10 | NES s 88: by agreement, not unreasonably refused; no award notice period. 4 weeks is AEL's *preferred* notice, so shorter goes to the senior approver rather than being refused |
| Personal/Sick | Entitlement, AllowPastDates, NeedsAttachment | NES s 107: notice as soon as practicable; evidence when there's a reason (Company Manual 6.14.3) |
| Carer's | Entitlement, AllowPastDates, NeedsAttachment | as sick |
| Compassionate | Entitlement, AllowPastDates | evidence only if the supervisor asks |
| Family & Domestic Violence | Entitlement, Confidential, AllowPastDates | NES s 106J–L, kept confidential |
| Community Service | Entitlement | NES s 110 |
| Long Service | SeniorApproval, MinNoticeDays 28 | state LSL Acts |
| Parental | SeniorApproval, **MinNoticeDays 70** | NES s 74: 10 weeks' written notice |
| Unpaid | SeniorApproval | discretionary |
| Time in Lieu | MinNoticeDays 7, MaxDaysBeforeSenior 5 | award cl 32.8, by agreement |

The agreements are Award Free; if the Manufacturing Award (MA000010) applies it adds no
employee notice rule — only **shutdowns** (cl 34.7: 28 days' written notice) and directing
**excess leave** (cl 34.9–34.11: over 8 weeks, starting 8 weeks–12 months out, at least a
week, leaving 6 weeks). Those are supervisor-booked leave, which the app doesn't do yet.

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

## Changing the rules

The rules above are AEL's **current** baseline (October 2026), not fixed in the code.
They can change — most are a number or a tick in SharePoint and take effect within a
minute, no redeploy:

| Rule | Where to change it |
|---|---|
| Notice needed (e.g. annual 4 weeks, parental 10 weeks) | LeaveTypes **MinNoticeDays** |
| Longest request before the senior approver (annual 10 days) | LeaveTypes **MaxDaysBeforeSenior** |
| Types that always need the senior approver | LeaveTypes **SeniorApproval** |
| When a certificate is needed | LeaveTypes **NeedsAttachment**, **CertificateAfterDays**, **CertificateNextToDayOff** |
| Backdating, contractors, confidential, retiring a type | LeaveTypes **AllowPastDates**, **Contractors**, **Confidential**, **Active** |
| How many can be off at once | Sites **MaxOnLeave** |
| Who the senior approver is | Sites **SeniorApprover** / env SENIOR_APPROVERS |

**Fixed in the code** (ask for a change): weekends and public holidays aren't counted;
a broken rule sends the request to the senior approver rather than refusing it;
entitlement types are acknowledged, not approved; nobody approves their own leave.

**When you change a rule, also update the wording** in each app's `guide.html` (the
"rules on asking for leave" boxes quote the current numbers) and `LeaveTypes_baseline.xlsx`.
The legal floor doesn't move: NES parental notice is 10 weeks, sick-leave evidence can
only be asked for with a reason, and annual leave can't be unreasonably refused.

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
