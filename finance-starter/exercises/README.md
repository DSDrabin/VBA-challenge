# Beginner exercises

Allow roughly 30–45 minutes per exercise. Record one result and one lesson in `notes/progress.md`.

| Step | Practical task | VBA skills | Done when |
|---|---|---|---|
| 1 | Count invoices and total invoice/paid amounts on Summary | Variables, Cells, last-row detection, For loop | 6 invoices; invoiced $9,200; paid $2,400 |
| 2 | Calculate outstanding = invoice amount − paid amount; flag unpaid invoices | Currency, If, writing cells | Outstanding total $6,800; 5 invoices have a balance |
| 3 | Calculate days overdue as at 30 June 2026 and assign ageing buckets | DateSerial, DateDiff, Select Case | Bucket balances match the checks below |
| 4 | Summarise outstanding balances by customer | Nested loops or a collection, summary tables, formatting | Alpha $800; Beta $2,500; Gamma $3,500 |
| 5 | Compare Bank and GL by unique reference; list exceptions | Worksheet loops, lookup logic, Currency comparisons | R003 mismatch; R004 bank only; R005 GL only |

## Ageing rules
Only positive outstanding balances enter ageing. Set days overdue = max(0, reporting date − due date).
Use **Current** when due date is on or after the reporting date, then **1–30**, **31–60**, **61–90**, **91+** days.
Expected balances: Current $2,500; 1–30 $800; 31–60 $1,000; 61–90 $0; 91+ $2,500.
The paid invoice must not appear in the outstanding report.

## Reconciliation rules
Amounts are signed: receipts positive, payments negative. Validate unique references before matching; stop and report duplicates rather than silently matching them.
Match reference first, then compare amount and date. Use a one-cent amount tolerance; flag date differences separately.
R003 bank − GL = **−$50**. Bank net movement is $1,250; GL is $1,150; difference **$100**.
Check: −$50 (mismatch) − $50 (bank only) − (−$200) (GL only) = $100.
These are movement comparisons, not a full opening/closing balance reconciliation.

## Stretch task: revisit the original stock macro
Read the archived original README and test a ticker with one row and one with several rows. Check whether the final row's volume is included, the opening price is captured for a single-row ticker, and yearly price change is shown as an amount. For zero opening prices, ensure a prior ticker's percent change cannot carry forward. Keep fixes separate from these exercises.
