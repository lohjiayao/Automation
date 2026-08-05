# Power Automate — Build Steps

Three flows, all built the same way: a trigger, a **Run script** action calling one of the
scripts in `/Scripts`, a **Parse JSON** step to read the result, and a **Send an email (V2)**
step. This file lists the exact configuration for each.

Prerequisite for all three: the tracker (`Connection_Tracker.xlsx`) is uploaded to OneDrive
or SharePoint, and the three scripts are pasted into it under **Automate → New Script** in
Excel Online, saved under their exact names (`MarkProjectPaid`, `GetOverdueProjects`,
`GetWeeklySummary`).

---

## Flow 1 — Payment Watcher

**Type:** Automated cloud flow
**Trigger:** Office 365 Outlook — *When a new email arrives (V3)*
- Folder: `Inbox`
- Subject Filter: `CC Bill Paid`
- From: the Finance sending address, or leave blank if multiple people send it

**Step 2 — Html to text** (Content Conversion)
- Content: **Body** (from the trigger)

**Step 3 — Run script** (Excel Online Business)
- Location: OneDrive for Business (or SharePoint in production)
- File: the tracker
- Script: `MarkProjectPaid`
- Parameters:
  - `emailSubject` → **Subject** (from trigger)
  - `emailBody` → **The plain text content** (from step 2)
  - `emailSender` → **From** (from trigger)

**Step 4 — Parse JSON**
- Content: **result** (from step 3)
- Schema (generate from sample):
```json
{"updated":true,"action":"created","ref":"SIB-260123","refsFound":1,"customer":"X","officerEmail":"a@b.com","engineerEmail":"e@b.com","reminderCount":0}
```

**Step 5 — Condition:** `action` (from Parse JSON) **is equal to** `created`
- **If yes → Send an email (V2)**
  - To: `engineerEmail`
  - Subject: `Verify project @{body('Parse_JSON')?['ref']} — not in tracker`
  - Body: explains a payment arrived for a project not yet in the tracker, that it's
    been added and marked paid, and asks the engineer to confirm the reference number.
- **If no →** leave empty (a normal matched payment updates silently, no email needed)

---

## Flow 2 — Daily Watchdog

**Type:** Scheduled cloud flow
**Trigger:** Recurrence — every **1 Day**, at **08:00**, time zone set to your local zone

**Step 2 — Run script**
- Script: `GetOverdueProjects`
- No parameters

**Step 3 — Parse JSON**
- Content: **result**
- Schema (generate from sample):
```json
{"recipients":"a@b.com;b@b.com","count":3,"html":"<div></div>"}
```

**Step 4 — Condition:** `count` (fx tab) **is greater than** `0`
- **If yes → Send an email (V2)**
  - To: `recipients`
  - Subject: `Daily check — @{body('Parse_JSON')?['count']} to action`
  - Body: `html`
- **If no →** leave empty (nothing to chase today — no email sent)

No loop is used — the script itself groups everything into one consolidated email.

---

## Flow 3 — Weekly Summary

**Type:** Scheduled cloud flow
**Trigger:** Recurrence — every **1 Week**, on your chosen day (e.g. Monday), **08:00**

**Step 2 — Run script**
- Script: `GetWeeklySummary`
- No parameters

**Step 3 — Parse JSON**
- Content: **result**
- Schema (generate from sample):
```json
{"recipients":"a@b.com;b@b.com","needsAttention":3,"active":12,"html":"<div></div>"}
```

**Step 4 — Send an email (V2)** (no Condition needed — sends every week regardless)
- To: `recipients`
- Subject: `Weekly summary — @{body('Parse_JSON')?['needsAttention']} need attention`
- Body: `html`

---

## Common pitfalls

- **Use "Generate from sample" with example data, not a schema.** Pasting a schema into
  that box produces a doubly-nested, invalid schema.
- **Whichever script you point Run script at, Parse JSON's schema and the Send email
  fields must match its actual output shape.** Mixing fields from two different scripts
  causes `ValidationFailed` or "property cannot be selected" errors.
- **If you remove a loop (Apply to each), remove every `items('Apply_to_each')?[...]`
  reference too** — replace with `body('Parse_JSON')?[...]`, or the flow fails to save.
- **Edit a script in place and Save — never delete and recreate one with the same name.**
  Scripts are tracked by a hidden ID; recreating breaks every flow pointing at the old one
  ("Script not found").
- **Trim email values read from Settings** (the scripts already do this) — a stray line
  break in a Settings cell produces an invalid "email\n" value that Send an email rejects.
