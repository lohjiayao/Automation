# Power Automate Build Guide — Connection Charge Tracker Automation

Companion file: **SESCO_Connection_Tracker.xlsx** (the shared tracker both flows read/write).

What you are building:

- **Flow 1 — Payment Watcher:** SEB Finance's payment email arrives → project automatically marked **Paid** in the tracker → technician + supervisor notified.
- **Flow 2 — Daily Watchdog:** every morning, scan the tracker → remind technicians about stuck items → escalate to the engineer when too old. Paid projects get a tighter clock.

No changes to SAP or ECX. Everything runs inside Microsoft 365.

---

## Part 0 — Prerequisites (do once)

1. **Put the tracker somewhere Power Automate can reach.**
   A cloud flow cannot read a file on your PC or on a network drive — it must be in
   OneDrive or SharePoint.

   - **While building and demoing: your own OneDrive for Business.** Sign in at office.com
     with your work account → **OneDrive** → *Upload* → `SESCO_Connection_Tracker_v2.xlsx`.
     Private to you, nothing shared with the team, and still inside the company tenant.
     In the *Run script* action, set **Location = OneDrive for Business**.
   - **For production: the planning team's Teams channel → *Files*.** Only move it there once
     the team agrees, since that's when real project data starts going in.

   The supplied tracker contains only invented example rows — no real customer data — so the
   demo version is safe to work with.

2. **Check Power Automate opens for you.**
   Go to <https://make.powerautomate.com> and sign in with your SESCO account. If it loads, you're set. If blocked, ask IT — mention it's for an internal planning-team automation.

3. **Reference number formats — already confirmed from real data:**
   - **WBS no.**: `SIB-260123`, sometimes typed without the dash (`SIB260124`). Pattern: `SIB` + 6 digits, where the first two digits are the year (`25` = 2025, `26` = 2026).
   - **SQ no.**: `SQ5001`, `SQ5002`. Pattern: `SQ` + 4–5 digits.
   Script 1 already matches both and ignores dashes/case when comparing.
   Finance's subject line reliably contains the ref, e.g.
   `[For Action] CC Bill Paid but no Budget Assignment - SIB-260123 SIB-C-EXAMPLE QUARRY SDN BHD`

4. **Mailbox decision for Flow 1.** The trigger watches a mailbox the flow's creator owns. Realistic options:
   - **Demo/testing:** run it on *your own* mailbox and send yourself fake Finance emails. ← start here
   - **Production:** your supervisor creates an Outlook rule auto-forwarding Finance emails to a team/shared mailbox the flow watches, **or** supervisor builds Flow 1 in their account by following this same guide (15 min).

---

## Part 1 — Add the two Office Scripts (the "real code")

Office Scripts are TypeScript programs stored in Excel Online; Power Automate calls them. They do the regex work Power Automate can't do natively.

**How to add a script:** open `SESCO_Connection_Tracker.xlsx` in **Excel Online** (from Teams: *Open in browser*) → **Automate** tab → **New Script** → delete the placeholder code → paste → rename the script (name matters, flows pick it by name) → **Save script**.

### Script 1 — name it `MarkProjectPaid`

```typescript
function main(
  workbook: ExcelScript.Workbook,
  emailSubject: string,
  emailBody: string
): string {
  // Matches the real SESCO formats seen in Finance emails and the planning sheet:
  //   SIB-260123 / SIB260124 (WBS no.)   and   SQ5001 / SQ5002 (SQ no.)
  const refPattern = /\b(SIB-?\d{6}|SQ-?\d{4,5})\b/i;

  // Ignore dashes/case/spaces when comparing ("SIB260124" == "SIB-260124")
  const norm = (s: string) => s.toUpperCase().replace(/[-\s]/g, "");

  // Count every reference in the mail, so a multi-project email cannot be
  // silently half-handled (that would recreate the very problem we are fixing).
  const refPatternAll = /\b(SIB-?\d{6}|SQ-?\d{4,5})\b/gi;
  const seen: string[] = [];
  for (const m of ((emailSubject || "") + " " + (emailBody || "")).match(refPatternAll) || []) {
    if (seen.indexOf(norm(m)) < 0) seen.push(norm(m));
  }
  const refsFound = seen.length;

  // Look in the SUBJECT first — Finance puts the ref there reliably —
  // then fall back to the body.
  let source = "subject";
  let match = (emailSubject || "").match(refPattern);
  if (!match) {
    source = "body";
    match = (emailBody || "").match(refPattern);
  }
  if (!match) {
    return JSON.stringify({
      found: false, ref: "", updated: false, source: "none", refsFound: refsFound
    });
  }
  const ref = match[0].toUpperCase();
  const refKey = norm(ref);

  const table = workbook.getTable("ProjectTracker");
  const range = table.getRangeBetweenHeaderAndTotal();
  const values = range.getValues();

  // Column positions (0-based) in the tracker table:
  // 0=WBS/SQ No, 1=Customer, 3=Planning Officer, 4=Officer Email,
  // 11=Paid, 12=Paid Date, 13=Budget Assigned
  const today = Math.floor(
    (Date.now() - Date.UTC(1899, 11, 30)) / 86400000
  ); // Excel date serial number

  for (let i = 0; i < values.length; i++) {
    if (norm(String(values[i][0])) === refKey) {
      range.getCell(i, 11).setValue("Yes");
      range.getCell(i, 12).setValue(today);
      range.getCell(i, 12).setNumberFormatLocal("DD/MM/YYYY");
      return JSON.stringify({
        found: true,
        ref: ref,
        updated: true,
        source: source,
        refsFound: refsFound,
        officerEmail: String(values[i][4]),
        officer: String(values[i][3]),
        customer: String(values[i][1]),
        budgetAssigned: String(values[i][13])
      });
    }
  }
  // Ref number detected in email but no matching row in tracker
  return JSON.stringify({
    found: true, ref: ref, updated: false, source: source, refsFound: refsFound
  });
}
```

### Script 2 — name it `GetOverdueProjects`

```typescript
function main(workbook: ExcelScript.Workbook): string {
  const table = workbook.getTable("ProjectTracker");
  const values = table.getRangeBetweenHeaderAndTotal().getValues();

  // Thresholds from the Settings sheet (yellow cells)
  const st = workbook.getWorksheet("Settings");
  const paidChaseDays = Number(st.getRange("B2").getValue()); // paid & idle
  const escalateDays = Number(st.getRange("B4").getValue());  // days past SLA
  const defaultSla = Number(st.getRange("B5").getValue());    // when blank

  const today = Math.floor(
    (Date.now() - Date.UTC(1899, 11, 30)) / 86400000
  );

  interface Item {
    ref: string; customer: string; officer: string; email: string;
    status: string; days: number; overBy: number; level: string;
  }
  const out: Item[] = [];

  for (const row of values) {
    const ref = String(row[0]).trim();
    const status = String(row[5]);
    const lastChange = Number(row[6]);
    const slaDays = Number(row[7]) > 0 ? Number(row[7]) : defaultSla;
    const paid = String(row[11]) === "Yes";
    const budgetAssigned = String(row[13]) === "Yes";

    if (!ref || !lastChange) continue;
    if (status === "Completed") continue;
    // Charge issued and unpaid = waiting on the customer, not on us
    if (status === "Charge Issued" && !paid) continue;

    const days = today - lastChange;
    const overBy = days - slaDays;

    let level = "";
    if (paid && !budgetAssigned) level = "ASSIGN-BUDGET";
    else if (paid && days >= paidChaseDays) level = "URGENT-PAID";
    else if (overBy >= escalateDays) level = "ESCALATE";
    else if (overBy > 0) level = "OVERDUE";
    if (!level) continue; // "due soon" shows in the sheet but sends no email

    out.push({
      ref: ref,
      customer: String(row[1]),
      officer: String(row[3]),
      email: String(row[4]),
      status: status,
      days: days,
      overBy: overBy,
      level: level
    });
  }
  return JSON.stringify(out);
}
```

---

## Part 2 — Flow 1: Payment Watcher

At <https://make.powerautomate.com>: **Create → Automated cloud flow** → name it `Payment Watcher` → trigger: **When a new email arrives (V2)** (Office 365 Outlook).

Build these steps in order:

1. **Trigger — When a new email arrives (V2)**
   - Folder: `Inbox`
   - *Show advanced options* → **Subject Filter:** `CC Bill Paid`
     Real Finance subjects look like:
     `[For Action] CC Bill Paid but no Budget Assignment - SIB-260123 SIB-C-EXAMPLE QUARRY SDN BHD`
     Filter on the subject, **not** the sender — senders vary, the subject phrase doesn't.
   - **From:** `financial@sarawakenergy.com` once Finance confirms a single sending address.
     If different Finance colleagues send it from their own accounts, leave **From** empty and
     rely on the subject filter alone — the wording is consistent, the sender is not.
   - For testing, leave **From** empty so you can trigger the flow by emailing yourself.

2. **Action — Html to text** (Content Conversion)
   - Content: dynamic content **Body** from the trigger.
   - (Finance emails are HTML; the script wants plain text. Note the SAP "Data Entry View" table is often a pasted *image*, so the body may contain little usable text — this is why the subject line is the primary source.)

3. **Action — Run script** (Excel Online (Business))
   - Location / Library / File: browse to `SESCO_Connection_Tracker.xlsx` on the team SharePoint.
   - Script: `MarkProjectPaid`
   - emailSubject: dynamic content **Subject** (from the trigger).
   - emailBody: dynamic content **The plain text content** (output of step 2).

4. **Action — Parse JSON** (Data Operation)
   - Content: dynamic content **Result** from Run script.
   - Schema → *Generate from sample* → paste:
     ```json
     {"found":true,"ref":"SIB-260123","updated":true,"source":"subject","refsFound":1,"officerEmail":"technician@sarawakenergy.com","officer":"ABC","customer":"Project","budgetAssigned":"No"}
     ```

5. **Action — Condition**
   - `updated` (from Parse JSON) **is equal to** `true`

   **If yes → Action — Send an email (V2)**
   - To: `officerEmail` (dynamic)
   - Subject: `PAID — @{body('Parse_JSON')?['ref']} needs budget assignment`
   - Body:
     ```
     Customer payment received for @{body('Parse_JSON')?['ref']}
     (@{body('Parse_JSON')?['customer']}).

     The tracker has been updated to Paid = Yes. Until Budget Assigned is set to Yes,
     this project shows as ACTION: ASSIGN BUDGET and will be chased daily.

     (Automated message from the Planning Tracker)
     ```
   - *(Optional extra: "Post message in a chat or channel" to your team's Teams channel with the same text.)*

   **If no → Action — Send an email (V2)** (to the supervisor)
   - Subject: `Payment email received — project not matched in tracker`
   - Body:
     ```
     A payment email arrived but the project could not be matched automatically
     (detected ref: "@{body('Parse_JSON')?['ref']}").
     Please update the tracker manually. Original email: @{triggerOutputs()?['body/subject']}
     ```

6. **Action — Condition** (safety net for batched emails)
   - `refsFound` (from Parse JSON) **is greater than** `1`
   - **If yes → Send an email (V2)** to the supervisor:
     `Payment email listed @{body('Parse_JSON')?['refsFound']} projects — only
     @{body('Parse_JSON')?['ref']} was updated automatically. Please check the rest manually.`
   - Without this, a single email covering several projects would mark one and silently
     leave the others stalled — exactly the failure this project exists to prevent.

6. **Save**, then **Test** → send yourself an email with the subject
   `[For Action] CC Bill Paid but no Budget Assignment - SIB-260123 TEST`
   (put `SIB-260123` in an example tracker row first) → within ~1 minute that row should
   show Paid = Yes and Priority = URGENT (PAID), and you should receive the notification email.

---

## Part 3 — Flow 2: Daily Watchdog

**Create → Scheduled cloud flow** → name `Daily Watchdog` → repeat **every 1 Day**, set the time (e.g. 08:00, time zone Kuala Lumpur).

1. **Action — Run script**: same Excel file, script `GetOverdueProjects`.

2. **Action — Parse JSON**
   - Content: **Result** from Run script.
   - Schema → *Generate from sample* → paste:
     ```json
     [{"ref":"SIB-260001","customer":"X","officer":"ABC","email":"a@b.com","status":"S","days":5,"overBy":2,"level":"OVERDUE"}]
     ```

3. **Action — Apply to each** — select output **Body** of Parse JSON. Inside the loop, a
   **Switch** on `level` (Data Operation → Switch) with four cases:

   | Case | Send an email (V2) to | Subject |
   |---|---|---|
   | `ASSIGN-BUDGET` | `email` , **Cc supervisor** | `[Action] @{items('Apply_to_each')?['ref']} — paid, budget not assigned (@{items('Apply_to_each')?['days']} days)` |
   | `URGENT-PAID` | `email` , **Cc supervisor** | `[Urgent] Paid project @{items('Apply_to_each')?['ref']} idle @{items('Apply_to_each')?['days']} days` |
   | `ESCALATE` | **engineer**, Cc `email` | `[Escalation] @{items('Apply_to_each')?['ref']} is @{items('Apply_to_each')?['overBy']} days past SLA` |
   | `OVERDUE` | `email` only | `[Reminder] @{items('Apply_to_each')?['ref']} is past its SLA date` |

   In each body include `customer`, `status`, `days` and `overBy` so the reader can act
   without opening the file. Keep the ASSIGN-BUDGET wording explicit — it maps to Finance's
   own phrase *"CC Bill Paid but no Budget Assignment"*, so recipients recognise it instantly.

4. **Save**, then **Test → Manually**. With the example rows dated in the past, you should immediately receive test reminder emails.

---

## Part 4 — Rollout checklist

- [ ] Replace the 3 example rows with real projects (keep formulas in Days Pending / Priority).
- [x] ~~Confirm real reference number format~~ — confirmed: `SIB-######` / `SQ####`.
- [ ] Check other Finance subject lines: is `CC Bill Paid` the wording every time, or do other payment-notification subjects exist? Widen the subject filter if so.
- [ ] Agree thresholds with supervisor → set them on the Settings sheet (no flow changes needed).
- [ ] Confirm what the 10 / 12 / 14 day figures represent, and whether they are calendar or working days.
- [ ] Confirm what STA(P) stands for (assumed: the reviewing party Planning submits to — nothing depends on it).
- [ ] Decide production mailbox (supervisor's rule → shared mailbox, or supervisor owns Flow 1).
- [ ] Put engineer's real email in Flow 2's escalate branch.
- [ ] Let it run in parallel with the current manual process for 1–2 weeks before relying on it.

## Design decisions (and why)

- **SLA lives per project, not in the code.** The tracker's `SLA Days` column mirrors the team
  sheet's `10/12 days`. Whoever keys the project enters 10, 12 or 14; every calculation follows
  from it. No one has to remember a global rule, and changing practice needs no flow edits.
- **Subject line, not sender.** Finance senders vary; the phrase `CC Bill Paid` does not.
- **Budget Assigned is tracked separately from Paid.** Finance's emails are specifically about
  payments landing with *no budget assignment*, so "paid" alone is not the finish line —
  the tracker chases until budget assignment is confirmed.
- **Separate file, not an edit of the live team sheet.** The team sheet already carries `#REF!`
  errors and mixed date formats; an intern-built automation should not be the thing that
  breaks it. Column names mirror the team's vocabulary so merging later is a small step.

## Part 4b — Capturing rejections (optional precision layer)

Flow 2 already protects against a forgotten rejection *without knowing one happened* — it
notices the project stopped moving. This part is about recording the rejection itself, so the
tracker shows **why** something stalled and counts rework rounds.

There are two ways to feed it, and **both use the same script**, so nothing is wasted if the
answer about SAP emails changes later.

### Script 3 — name it `MarkProjectRejected`

```typescript
function main(
  workbook: ExcelScript.Workbook,
  projectRef: string,
  reason: string
): string {
  const norm = (s: string) => s.toUpperCase().replace(/[-\s]/g, "");
  const refPattern = /\b(SIB-?\d{6}|SQ-?\d{4,5})\b/i;

  const m = (projectRef || "").match(refPattern);
  if (!m) return JSON.stringify({ updated: false, ref: "", note: "no valid reference" });
  const refKey = norm(m[0]);

  const table = workbook.getTable("ProjectTracker");
  const range = table.getRangeBetweenHeaderAndTotal();
  const values = range.getValues();

  const today = Math.floor((Date.now() - Date.UTC(1899, 11, 30)) / 86400000);
  const nextRound = (cur: string): string => {
    if (cur === "1st") return "2nd";
    if (cur === "2nd") return "3rd";
    if (cur === "3rd" || cur === "4th+") return "4th+";
    return "2nd"; // blank or unexpected -> first rejection means round 2
  };

  for (let i = 0; i < values.length; i++) {
    if (norm(String(values[i][0])) === refKey) {
      range.getCell(i, 5).setValue("Rejected - Pending Resubmission"); // Status
      range.getCell(i, 6).setValue(today);                             // Last Status Change
      range.getCell(i, 6).setNumberFormatLocal("DD/MM/YYYY");
      range.getCell(i, 10).setValue(nextRound(String(values[i][10])));  // Submission Round
      const old = String(values[i][15] || "");
      range.getCell(i, 15).setValue(
        (old ? old + " | " : "") + "Rejected: " + (reason || "no reason given")
      );
      return JSON.stringify({
        updated: true,
        ref: m[0].toUpperCase(),
        officerEmail: String(values[i][4]),
        officer: String(values[i][3]),
        customer: String(values[i][1]),
        round: nextRound(String(values[i][10]))
      });
    }
  }
  return JSON.stringify({ updated: false, ref: m[0].toUpperCase(), note: "no matching row" });
}
```

Resetting **Last Status Change** to today is the important part: the SLA clock restarts, and
Flow 2 begins chasing from that moment.

### Route A — SAP (or the engineer) sends an email

*Only possible if rejections generate an email.* Same shape as Flow 1:

1. Trigger **When a new email arrives (V3)**, Subject Filter set to whatever the rejection
   mail contains (e.g. `Rejected`).
2. **Html to text** → Body.
3. **Run script** → `MarkProjectRejected`, `projectRef` = **Subject**, `reason` = the plain
   text content.
4. **Parse JSON** → **Condition** on `updated` → notify the officer.

### Route B — a Microsoft Form *(works today, needs nothing from SAP)*

1. At **forms.office.com**, create a form "Log a Cost Rejection" with two questions:
   *Project No* (text, required) and *Reason for rejection* (text).
2. New **Automated cloud flow** → trigger **When a new response is submitted** (Microsoft Forms).
3. Action **Get response details**.
4. **Run script** → `MarkProjectRejected`, passing the two answers.
5. **Parse JSON** → **Condition** → email the assigned officer.

Twenty seconds of typing for the engineer, and the rejection is permanently recorded with a
restarted SLA clock. Honest limitation: it depends on the engineer remembering to fill it in —
which is the habit that fails today. That is exactly why Flow 2 remains the backbone, and this
is a supplement, never the main mechanism.

## Part 5 — Handover: getting this onto someone else's account

**The rule that governs everything here:** a flow watches the mailbox of the account whose
**connection** it uses. Building it in your account means it watches *your* inbox — so it will
never see Finance's emails to your supervisor, no matter who the flow is shared with.

**And the one that matters most:** when your internship ends and your account is disabled,
any flow you own **stops running**. The production version must not be owned by you.

### Option 1 — Shared mailbox *(best; recommended for production)*

Ask IT (via Assyst) for a shared mailbox, e.g. `planning.cc@sarawakenergy.com`, and have
Finance send there — or have your supervisor add an Outlook rule forwarding `CC Bill Paid`
mail to it.

Then in the flow, swap the trigger to **"When a new email arrives in a shared mailbox (V2)"**
and enter the shared mailbox address. Everything downstream stays identical.

Why this is the right answer: the flow depends on a *mailbox*, not a *person*. It survives
your internship ending, your supervisor going on leave, and staff changes.

### Option 2 — Export and import into the supervisor's account

Hands over a working flow without rebuilding it by hand.

1. **You:** *My flows* → the flow's **⋯** → **Export** → **Package (.zip)** → download.
2. **Supervisor:** *My flows* → **Import** → upload the .zip.
3. During import, next to each connection click **Select during import** and pick
   **their own** Office 365 Outlook and Excel connections.
4. Import, then **turn the flow on**.

From then on it watches *their* inbox and they own it. Note their Excel connection also
needs edit access to the tracker on SharePoint.

### Option 3 — Forwarding rule *(quick interim, for demos)*

Your supervisor creates an Outlook rule: subject contains `CC Bill Paid` → forward to you.
Your existing flow then sees the mail unchanged. Fine for proving the concept; not a
production answer, since it still dies when your account does. Some tenants also block
auto-forwarding by policy.

### Recommended path

Demo in your own account → export/import to your supervisor once it works → move to a
shared mailbox before your internship ends.

## Appendix — Why not AI?

Power Automate does offer AI for reading email content (**AI Builder**: text recognition/OCR,
document processing, entity extraction, GPT prompts). It was considered and deliberately not
used:

- **Licensing.** AI Builder needs premium credits — a procurement conversation, not a
  standard M365 entitlement. Every connector in this solution is standard.
- **Reliability.** With Finance's agreed subject format, a pattern match on `SIB-######` /
  `SQ####` is deterministic — it either matches or it doesn't. An AI extraction can be
  confidently wrong, and a wrongly-matched project number silently updates the wrong row.
- **Explainability.** When a run misbehaves, the run history shows exactly what was matched.

**The one case where AI would genuinely help:** reading the pasted SAP "Data Entry View"
*screenshot* to recover the amount and payment date, since an image cannot be parsed as text.
That need disappears if Finance includes the plain-text block requested in
`Finance_Email_Format_Request.md` — which is free, and more reliable than OCR.

## Troubleshooting quick answers

- **"Run script" can't find the file** → the tracker must be on SharePoint/OneDrive, not a network drive, and you need edit access.
- **Flow 1 never triggers** → trigger only watches the mailbox of whoever created the flow; check the From/Subject filters aren't too strict.
- **Paid Date shows a number like 46234** → select the column → format as Date (the script sets DD/MM/YYYY on the cell it writes, but a manually cleared format can override).
- **New rows show blank Days Pending/Priority** → copy those two formulas down from the row above (Excel tables usually auto-fill them).
- **Script 1 matches the wrong number** (e.g. a phone number) → tighten `refPattern`, e.g. `/\b400\d{5}\b/` if all projects start with 400.
