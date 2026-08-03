# Requested Email Format — Connection Charge Payment Notifications

**To:** SEB Finance (`financial@sarawakenergy.com`)
**From:** Planning Team
**Purpose:** allow payment notifications to update the planning tracker automatically,
so paid projects are never missed while waiting for budget assignment.

Nothing about Finance's process changes — this is only about keeping the wording and
layout consistent so the information can be read automatically.

---

## Requested subject line

```
[For Action] CC Bill Paid - <Project No> - <Customer Name>
```

Example:

```
[For Action] CC Bill Paid - SIB-260123 - EXAMPLE QUARRY SDN BHD
```

Two things matter here:

1. **Keep the phrase `CC Bill Paid`** — this is what identifies the email.
2. **Put the project number immediately after it**, in the standard format
   (`SIB-260123` or `SQ5001`).

## Requested body block

Please include these lines as **plain text** near the top of the email. The existing SAP
"Data Entry View" screenshot can still be attached below it as usual.

```
Project No   : SIB-260123
Customer     : EXAMPLE QUARRY SDN BHD
Payment Date : 24.07.2026
Amount (MYR) : 10,623.00
Budget Assignment Required : Yes
```

**Why plain text matters:** the SAP table is currently pasted as an *image*. Images cannot
be read automatically — the information is visible to a person but invisible to the system.
Typing these five lines takes a few seconds and makes the whole notification machine-readable.

## One project per email

Where practical, please send **one email per project**. If several projects must be
combined, the planning team will still be alerted — but each project has to be checked
manually, which is what this change is trying to avoid.

---

## What Planning does with it

The email is picked up automatically and the project is marked **Paid** in the planning
tracker, with the assigned officer notified. The project then stays flagged as
**"Action: Assign Budget"** until the budget assignment is completed — so a paid customer's
project cannot sit unnoticed.

No reply to the email is needed. Finance's existing recipients and process stay exactly
as they are today.
