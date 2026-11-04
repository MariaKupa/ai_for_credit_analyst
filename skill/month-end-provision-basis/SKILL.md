---
name: month-end-provision-basis
description: Produces the month-end provision basis – exposure and ECL per product, currency and IFRS 9 stage – compares it with the previous month-end, saves an Excel file and sends a notification with a short AI-written comment. Use when the scheduled "Month-end provision basis" automation runs, or when someone asks for the provision basis, the hand-over file to finance, or a month-over-month exposure/ECL comparison.
---

# Month-end provision basis

Before this skill, an analyst built the file by hand on the first business day of each month. The skill does it as follows:

1. Sums `exposure` and `ecl` per `product`, `currency` and `stage` for the latest month-end.
2. Checks that the data is fresh and compares it with the previous month-end and with last month's delivered file.
3. Saves `provision_basis_YYYY-MM-DD.xlsx`, with the output and the comparisons, automatically to the given path (`--output-dir` or the `PROVISION_OUTPUT_DIR` environment variable).
4. Sends a notification with the checks, a short comment and a table per product and stage.

**Division of labour:** everything deterministic (SQL, checks, Excel, comparison, notification) is done by `scripts/provision_basis.py`. The agent only runs the script, reads the result and writes the comment. **The agent never calculates or changes numbers.**

## Agent workflow

Run from the demo root:

```powershell
python skill\month-end-provision-basis\scripts\provision_basis.py build
```

For a manual run, add `--manual` (no business-day rule). The script prints `STATUS=`, `REPORT_FILE=`, `SUMMARY_JSON=` and `NOTIFY=` lines. Act on the exit code:

| Exit | STATUS | What the agent does |
|---|---|---|
| 0 | `done` | Read `SUMMARY_JSON`. Write the comment (rules below) to `comment.md` next to summary.json. Run `notify --summary <SUMMARY_JSON> --comment-file <comment.md>`. Reply briefly: file name, whether the notification was sent, and the comment. |
| 2 | `not_ready` | The data is not rebuilt after month-end or has no rows for the report date. Nothing was saved. Say what is missing. Do **not** use `--force` unless the user explicitly asks. |
| 3 | `skipped` | Nothing to do (not business day 1–3, or the file already exists). Reply in one sentence. |
| 1 | `error` | Report the error and suggest a fix. Do not change code or data. |

## Rules for the comment

- 3–6 bullet points (`- `), factual and short. No headings or tables; the table is added by the script.
- Use **only** facts from summary.json: `totals_by_product`, `comparison`, `notable_changes`, `restatements`, `readiness` and `warnings`. **Never guess causes that are not visible in the data.**
- If products are missing (`readiness.missing_details`), say so in the first bullet: which products, last month's exposure and ECL, last date with rows, and that the file was saved without them.
- Then describe the products that changed most, judged against each product's own size. Never compare amounts in different currencies.
- Then name the stages behind the changes (`notable_changes`), new or removed products, restated values for the previous month-end (`restatements.rows`) and any warnings.
- Amounts in millions with the currency code (`1,598.0 MEUR`), percentages as `+5.3 %`.

## Flags

| Flag | Meaning |
|---|---|
| `build --manual` | Run now, regardless of business day. Skips if the month's file already exists. |
| `--date YYYY-MM-DD` | Other report date (a month-end). |
| `--overwrite` | Replace an existing file for the same month. |
| `--force` | Save even if data checks fail (a warning is added). |
| `--dry-run` | Write to `%TEMP%\provision-basis\` and send nothing. |
| `notify --summary <file> [--comment-file <file>] [--dry-run \| --test]` | Build and send (or resend) the notification. |

Notifications are always saved as `notification.md` next to summary.json. They are posted as JSON `{"text": ...}` only if `PROVISION_WEBHOOK_URL` is set. Keep the URL in an environment variable, never in the repo or the chat.

## Business rules in the script

- **Report date:** the last day of the previous month. It is compared with the month-end before, using the database's current values for both.
- **Readiness:** the table must have been rebuilt after month-end and have rows for the report date, otherwise no file is saved. Products missing compared with last month do not block the run; they are flagged.
- **Notable change:** |Δ %| ≥ 5 % **and** |Δ| ≥ 0.5 % of the currency total, plus new and removed rows.
- **Restatements:** last month's delivered file is compared with today's database values; differences > 1 (local currency) are reported.
- **Reconciliation:** the Excel file is read back and compared with the database before it is saved to the given path.

## Guardrails

- The agent may write the required `comment.md` next to `summary.json`. Apart from that commentary file, never edit files, commit or open pull requests during an automation run. The script remains responsible for writing the summary, Excel output and notification.
- Run `build` at most once per run; never use `--overwrite` or `--force` unless asked.
- Only aggregated amounts per product, currency and stage are read by the AI model, never customer-level data.
