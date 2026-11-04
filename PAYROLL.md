# Payroll page

**https://au-approvals.ael.co/payroll** is the page payroll uses to run the pay week. It's part of the approvals app, with the same Microsoft 365 sign-in, and it can sit on a SharePoint page.

## What payroll can do

- **See every site at a glance.** The home page shows each site's last finished pay week:
  - approved hours, and how many people are approved;
  - who's still waiting for approval, has been sent back, or hasn't submitted;
  - leave hours;
  - whether the week has been processed.
- **Open a site's week.** You see everyone at that site, with hours per day, total, allowances and leave, plus a status for each person:
  - Approved;
  - Waiting for approval;
  - Waiting for senior approver;
  - Sent back;
  - Not submitted.

  Click a person to see their lines.
- **Chase what's outstanding.** The Outstanding table lists everyone who isn't approved. Each person has an **Email** button, and **Email all outstanding** writes one email to all of them with everyone in Bcc. Both open a ready-written email in Outlook, and nothing is sent until you send it.
- **Correct approved lines.** You can:
  - change hours, work order or comment on a line;
  - remove a line;
  - add a line of hours or an allowance on any day of the week.

  Write a reason, then **Save changes**. Each change is stamped on the line's Notes as "Payroll changed … (by you on date — reason)" and listed under **Payroll changes** for that person. A removed line is noted on the person's other lines. Only approved lines can be changed here: waiting and sent-back lines stay with the supervisor and the worker.
- **See what allowances to pay (Allowances view).** In a site's week, switch from **People** to **Allowances**. You get:
  - a card for each allowance with the site's total, coloured by kind:
    - **per shift / day**: daily allowances from the Allowances list, e.g. Night Shift, Weekend, Public holiday;
    - **per job**: the call-out ticks on hours lines;
    - **weekly**: e.g. On call;
  - the site's **Shifts worked** totals (from each day's Shift worked);
  - a table with one row per person and one column per allowance. Each cell shows how many, plus which days ("Mon · Wed · Thu ×2"); a weekly allowance shows ✓. There's a total row at the bottom.

  Only approved lines are counted. Tick **Include not yet approved** to see the rest as "+n waiting". Processed people are grey. Click a name to jump to their lines. **Print** prints the table, and the Excel download has the same table on an **Allowances** sheet (the PDF has a short list). An allowance's kind comes from its **Frequency** in the Allowances list.
- **Download Excel or PDF.** These hold the week's approved hours, in the same layout as the automatic payroll email. Approved leave is included too: a **Leave** sheet in Excel, and a Leave section in the PDF.
- **Process as you go.** You can process at three levels:
  - **Each line** has a tick box: tick it when you've processed that line.
  - **Process all** on a person processes every approved line they have that week.
  - **Process all approved** at the top processes the whole site.

  Processed lines and people turn grey, and each line shows who processed it and when. To change a processed line, untick it, or click **✓ Processed** on the person to undo theirs. The filter chips (**To process**, **Processed**, **Outstanding**) and the name search narrow the list. Lines approved later arrive unticked, so they show up under **To process**. Processing is saved on each line's Notes ("Processed for payroll by … on …").

- **Look up any employee (Employees tab).** Search for anyone on the Workers or Contractors lists, including inactive people. The **site** drop-down narrows the search to people whose home site it is. With a site picked and nobody searched for yet, everyone at that site is shown as cards to click. The page remembers your site for next time. Then pick a period:
  - the last 4 or 12 weeks, 6 months or 12 months;
  - **One pay week**: any date, and you get the pay week it falls in;
  - **Between dates**, up to a year.

  You see their details (type, home site, email, contract hours), totals and leave, then each pay week, newest first. Open a week to see every line:
  - site, work order or allowance, and hours;
  - call-out tags;
  - shift worked against the rostered shift;
  - comment;
  - status, who approved it and when, any send-back reason, and who processed it.

  Each week also shows when the timesheet was first saved and last changed, and its full history of notes. **Print** prints the lot. **Open … pay week** jumps to that site's week to process or correct it. The address in the browser links straight back to the same person and period.

## Setup

1. **Give payroll access.** In the timesheet **Settings** list, add a row:
   - **Title** `PayrollAdmins`
   - **Value**: the payroll people's emails, separated by `;`

   Senior approvers (Netlify `SENIOR_APPROVERS`) can use it too. Netlify env `PAYROLL_ADMINS` works as well.
2. **Put it on SharePoint**, either way:
   - **Embed it in a page.**
     - First allow the domain. In **Site settings → HTML Field Security**, add `au-approvals.ael.co`.
     - Then edit a page, add an **Embed** web part and paste:
       `<iframe src="https://au-approvals.ael.co/payroll" width="100%" height="1100" style="border:0"></iframe>`
     - Inside SharePoint, sign-in opens in a small pop-up window. Allow pop-ups if the browser asks.
   - **Or add a link.** Use a Quick Links web part pointing to https://au-approvals.ael.co/payroll, which opens it full screen.

No new SharePoint lists or columns are needed.

## Who's expected at a site

"Not submitted" comes from:
- the **Workers** list (staff, by HomeSite / Site);
- active **Contractors** (by Site).

People are matched to their timesheets by name, or by the first-initial.surname@ael.co email. If someone shows twice, check that their name matches in Workers / Contractors.

Code: `payroll.html`, `netlify/functions/payroll.js`. The Excel and PDF come from `report.js`.
