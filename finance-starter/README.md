# VBA finance starter

Practice Excel automation using fictional invoice and reconciliation data. This is separate from the original stock challenge at the repository root.

## Start
1. Use desktop Excel with VBA support. Open a new workbook and save it locally as `finance_practice.xlsm`.
2. Import `data/invoices.csv` through Data > From Text/CSV into a sheet named `Invoices`; use ISO dates and numeric amount columns. Rename the other sheets `Summary`, `Bank`, and `GL`.
3. Import the other CSV files into Bank and GL. In the VBA editor, import `modules/FinanceStarter.bas`.
4. Work through [the exercises](exercises/README.md), one at a time. Export your modules after each exercise.

## Folders
| Folder | Purpose |
|---|---|
| `data/` | Fictional, unchanged input CSVs |
| `modules/` | Exported VBA source; easier to review than workbook binaries |
| `exercises/` | Tasks and expected results |
| `workbooks/` | Local macro-enabled practice workbooks |
| `outputs/` | Generated summaries and screenshots |
| `notes/` | What you learned and problems to revisit |

Use `Option Explicit`, `Long` row counters, `Currency` for money, and fully qualified worksheet references. Write results to Summary; keep inputs unchanged. Re-running a macro should replace its previous output.

All amounts are AUD; reporting date is **30 June 2026**. These are learning datasets, not company records.
