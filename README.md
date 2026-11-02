# Month-end IFRS 9 output with an AI agent

How I, as a credit risk analyst, automated the monthly IFRS 9 output with an AI skill. The script extracts and checks the data, compares it with last month and saves the Excel file. The AI writes a short comment on what changed.

📖 **The story behind it:** [Use AI to automate the IFRS9 output](https://mariakupa.se/posts/month-end-provision-basis)

> Shared for inspiration. All products, amounts and events are synthetic. No company data, systems or names are used. © 2026, all rights reserved.

## What the skill does

Every month-end, it:

1. Checks that the data is refreshed after month-end and that all products are present.
2. Sums exposure and expected credit loss (ECL) per product, currency and IFRS 9 stage.
3. Compares with last month, in total and per stage, and flags the large changes.
4. Compares last month's delivered file with the database today, to catch history that has changed (restatements).
5. Saves an Excel file with the output and the comparisons, after reconciling it with the database.
6. Sends a notification with the checks, a short AI-written comment and a table per product and stage.

## The design: AI never touches the numbers

| Step | Who does it |
|---|---|
| Extract, check, compare, write and reconcile the Excel file | Script (`provision_basis.py`) |
| Read the results and write 3–6 bullet points on what changed | AI agent (rules in `SKILL.md`) |
| Decide whether to override a failed check | Human |

The script saves all results in `summary.json`. That file is the only thing the AI reads, so the comment can only be based on facts the script has calculated.

The script tells the agent what happened with an exit code: `0` done, `1` error, `2` data not ready, `3` nothing to do. `SKILL.md` says exactly how the agent should react to each one.

## Example notification

From the synthetic demo, October 2026. Product A's September numbers were corrected after delivery, Product B has a stage 2 migration, and Product C is missing:

```text
✔ Provision basis 2026-10-31 saved – Product C missing

Checks
- Table rebuilt 2026-11-01T06:05:00 ✔
- 6 rows for 2026-10-31 ✔
- ⚠ Product C missing (last month 12.7 MEUR exposure, 0.2 MEUR ECL, last rows 2026-09-30)
- ⚠ 1 restated row(s) vs provision_basis_2026-09-30.xlsx

Comment (written by the AI)
- Product C is missing for 2026-10-31. Last month it had 12.7 MEUR exposure and 0.2 MEUR ECL. The file was saved without it.
- Product B ECL rose +8.6 % to 7.0 MUSD while exposure fell -0.9 % to 315.6 MUSD.
- The increase comes from stage 2: exposure +9.2 MUSD (+36.3 %) and ECL +0.6 MUSD (+31.3 %).
- Product A ECL fell -2.7 % to 40.6 MEUR, with exposure down -2.0 % to 1,598.0 MEUR.
- Restatement: stage 2 ECL for Product A at 2026-09-30 is now 0.18 MEUR higher than delivered (6.05 → 6.23 MEUR).

Exposure and ECL per product and stage
| Product   | Stage | Exposure     | Δ %    | ECL        | Δ %    |
| Product A | all   | 1,598.0 MEUR | -2.0 % | 40.6 MEUR  | -2.7 % |
| Product B | all   | 315.6 MUSD   | -0.9 % | 7.0 MUSD   | +8.6 % |
| ...       |       |              |        |            |        |
```

## What's in the repository

| File | What it is |
|---|---|
| `skill/month-end-provision-basis/SKILL.md` | The instructions for the AI agent: what to run, how to react to each result and the rules for the comment. |
| `skill/month-end-provision-basis/scripts/provision_basis.py` | The script that does all the work with numbers: SQL, checks, comparison, Excel file and notification. |
| `data/generate_synthetic_portfolio.py` | Creates a small fake database with three products for June–October 2026. It stands in for the real IFRS 9 model output. |

## Run the demo

You need Python 3.11 or newer. Open PowerShell in this folder and run:

```powershell
# 1. Install the Excel package
python -m pip install openpyxl

# 2. Create the fake data (refreshed on 1 November, the day after month-end)
python data\generate_synthetic_portfolio.py --rebuilt-at 2026-11-01T06:00:00

# 3. Deliver September's file, as if it was done last month
python skill\month-end-provision-basis\scripts\provision_basis.py build --manual --date 2026-09-30

# 4. Correct Product A's September numbers after delivery (a restatement)
python data\generate_synthetic_portfolio.py --apply-restatement --rebuilt-at 2026-11-01T06:05:00

# 5. Run October: checks, comparison, Excel file and summary.json
python skill\month-end-provision-basis\scripts\provision_basis.py build --manual --date 2026-10-31

# 6. Build the notification (saved to runs\2026-10-31\notification.md, nothing is sent)
python skill\month-end-provision-basis\scripts\provision_basis.py notify --summary runs\2026-10-31\summary.json --dry-run
```

What the flags mean:

- `--manual` runs now. Without it, the script only runs on business days 1–3, as in a scheduled automation.
- `--date` chooses the month-end.
- `--dry-run` only saves the notification. Without it, and with `PROVISION_WEBHOOK_URL` set, it's also posted, for example to Teams.

In a real run, the AI writes `runs\2026-10-31\comment.md` between step 5 and 6, and step 6 adds `--comment-file runs\2026-10-31\comment.md`. Without a comment, the script lists the largest changes itself.

The Excel file ends up in `output\`. Use `--output-dir` or the environment variable `PROVISION_OUTPUT_DIR` to save it somewhere else. To start over, delete `output`, `runs` and `data\portfolio.db`.

## How it runs with GitHub Copilot

The skill folder lives in `.github/skills/` in a repository. You can ask Copilot to *"produce the month-end provision basis"*, or run it as a scheduled automation on business days 1–3. Copilot reads `SKILL.md`, runs the script, writes the comment and sends the notification.
