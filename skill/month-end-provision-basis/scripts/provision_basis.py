"""Month-end provision basis: exposure and ECL per product, currency and IFRS 9 stage.

Everything deterministic happens here: SQL, data checks, the Excel file, the comparison with
the previous month-end and the notification. The AI agent only runs this script, reads
summary.json and writes a short comment (see ../SKILL.md).

Commands:
  build   extract, check, compare and save the Excel file + summary.json
  notify  build the notification from summary.json (+ optional comment) and send/save it

Machine-readable output lines: STATUS=, REPORT_FILE=, SUMMARY_JSON=, NOTIFY=
Exit codes: 0 done, 1 error, 2 not_ready, 3 skipped
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import shutil
import sqlite3
import sys
import tempfile
import urllib.request
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

ROOT = Path(__file__).resolve().parents[3]
DEFAULT_DB = ROOT / "data" / "portfolio.db"
DEFAULT_OUTPUT = ROOT / "output"
RUNS_DIR = ROOT / "runs"

CHANGE_PCT_THRESHOLD = 0.05                # |Δ %| threshold for a notable change
NOTABLE_SHARE_OF_TOTAL = 0.005    # ... and |Δ| must be at least 0.5 % of the currency total
RESTATE_TOLERANCE = 1.0       # in local currency
BUSINESS_DAYS = (1, 2, 3)         # scheduled runs only act on business day 1-3

EXTRACT_SQL = """SELECT value_date, product, currency, stage,
       SUM(exposure) AS exposure, SUM(ecl) AS ecl
FROM data_source
WHERE value_date = ?
GROUP BY value_date, product, currency, stage
ORDER BY product, currency, stage"""


class BasisError(Exception):
    pass


# ---------------------------------------------------------------- helpers
def previous_month_end(day: dt.date) -> dt.date:
    return day.replace(day=1) - dt.timedelta(days=1)


def business_day_of_month(day: dt.date) -> int | None:
    """1-based business day number (Mon-Fri, no holiday calendar in the demo)."""
    if day.weekday() >= 5:
        return None
    return sum(1 for d in range(1, day.day + 1) if day.replace(day=d).weekday() < 5)


def basis_filename(report_date: dt.date) -> str:
    return f"provision_basis_{report_date.isoformat()}.xlsx"


def pct(curr: float, prev: float) -> float | None:
    return None if not prev else (curr - prev) / abs(prev)


def round_floats(obj, key: str = ""):
    if isinstance(obj, float):
        return round(obj, 4 if key.endswith("pct") else 2)
    if isinstance(obj, dict):
        return {k: round_floats(v, k) for k, v in obj.items()}
    if isinstance(obj, list):
        return [round_floats(v, key) for v in obj]
    return obj


def format_millions(value: float | None, ccy: str = "") -> str:
    if value is None:
        return "–"
    return f"{value / 1e6:,.1f} M{ccy}"


def format_pct(value: float | None) -> str:
    return "–" if value is None else f"{value * 100:+.1f} %"


# ---------------------------------------------------------------- data
def connect(db: Path) -> sqlite3.Connection:
    if not db.exists():
        raise BasisError(f"Database not found: {db}. Run data/generate_synthetic_portfolio.py first.")
    conn = sqlite3.connect(db)
    conn.row_factory = sqlite3.Row
    return conn


def load_totals(conn, value_date: dt.date) -> dict:
    rows = conn.execute(EXTRACT_SQL, (value_date.isoformat(),)).fetchall()
    return {(r["product"], r["currency"], r["stage"]): {"exposure": r["exposure"], "ecl": r["ecl"]}
            for r in rows}


def data_checks(conn, report_date: dt.date, curr: dict, prev: dict) -> dict:
    rebuilt = conn.execute("SELECT rebuilt_at FROM load_log WHERE table_name = 'data_source'").fetchone()
    rebuilt_at = dt.datetime.fromisoformat(rebuilt[0]) if rebuilt else None
    curr_products = {k[0] for k in curr}
    prev_products = {k[0] for k in prev}
    missing = sorted(prev_products - curr_products)
    missing_details = []
    for product in missing:
        last = conn.execute("SELECT MAX(value_date) FROM data_source WHERE product = ?",
                            (product,)).fetchone()[0]
        rows = [v for k, v in prev.items() if k[0] == product]
        missing_details.append({
            "product": product,
            "currency": next(k[1] for k in prev if k[0] == product),
            "prev_exposure": sum(r["exposure"] for r in rows),
            "prev_ecl": sum(r["ecl"] for r in rows),
            "last_value_date": last,
        })
    rebuilt_ok = rebuilt_at is not None and rebuilt_at.date() > report_date
    return {
        "rebuilt_at": rebuilt_at.isoformat() if rebuilt_at else None,
        "rebuilt_after_month_end": rebuilt_ok,
        "row_count": len(curr),
        "ready": rebuilt_ok and len(curr) > 0,
        "missing_products": missing,
        "missing_details": missing_details,
        "new_products": sorted(curr_products - prev_products),
    }


def compare_months(curr: dict, prev: dict) -> list[dict]:
    rows = []
    for key in sorted(set(curr) | set(prev)):
        c, p = curr.get(key), prev.get(key)
        row = {"product": key[0], "currency": key[1], "stage": key[2],
               "status": "new" if p is None else "gone" if c is None else ""}
        for m in ("exposure", "ecl"):
            cv, pv = (c or {}).get(m, 0.0), (p or {}).get(m, 0.0)
            row |= {f"{m}_prev": pv, f"{m}_curr": cv, f"{m}_delta": cv - pv, f"{m}_pct": pct(cv, pv)}
        rows.append(row)
    return rows


def sum_by(comparison: list[dict], fields: tuple) -> list[dict]:
    out: dict[tuple, dict] = {}
    for r in comparison:
        key = tuple(r[f] for f in fields)
        agg = out.setdefault(key, {f: r[f] for f in fields} | {"exposure_prev": 0.0, "exposure_curr": 0.0,
                                                               "ecl_prev": 0.0, "ecl_curr": 0.0})
        for m in ("exposure_prev", "exposure_curr", "ecl_prev", "ecl_curr"):
            agg[m] += r[m]
    for agg in out.values():
        for m in ("exposure", "ecl"):
            agg[f"{m}_delta"] = agg[f"{m}_curr"] - agg[f"{m}_prev"]
            agg[f"{m}_pct"] = pct(agg[f"{m}_curr"], agg[f"{m}_prev"])
    return list(out.values())


def flag_changes(comparison: list[dict]) -> list[dict]:
    totals = {t["currency"]: t for t in sum_by(comparison, ("currency",))}
    out = []
    for r in comparison:
        for m in ("exposure", "ecl"):
            total = max(totals[r["currency"]][f"{m}_curr"], totals[r["currency"]][f"{m}_prev"])
            big = r[f"{m}_pct"] is not None and abs(r[f"{m}_pct"]) >= CHANGE_PCT_THRESHOLD \
                and abs(r[f"{m}_delta"]) >= NOTABLE_SHARE_OF_TOTAL * total
            if big or (r["status"] and r[f"{m}_delta"]):
                out.append({"product": r["product"], "currency": r["currency"], "stage": r["stage"],
                            "measure": m, "prev": r[f"{m}_prev"], "curr": r[f"{m}_curr"],
                            "delta": r[f"{m}_delta"], "pct": r[f"{m}_pct"], "status": r["status"]})
    return sorted(out, key=lambda x: -abs(x["delta"]) / max(totals[x["currency"]][f"{x['measure']}_curr"], 1))


def read_delivered_file(path: Path) -> dict:
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        out = {}
        for row in wb["Basis"].iter_rows(min_row=2, values_only=True):
            if row[1] and row[1] != "TOTAL":
                out[(row[1], row[2], int(row[3]))] = {"exposure": float(row[4]), "ecl": float(row[5])}
        return out
    finally:
        wb.close()


def find_restatements(output_dir: Path, prev_date: dt.date, db_prev: dict) -> dict:
    path = output_dir / basis_filename(prev_date)
    if not path.exists():
        return {"previous_file": None, "rows": []}
    delivered = read_delivered_file(path)
    rows = []
    for key in sorted(set(delivered) | set(db_prev)):
        for m in ("exposure", "ecl"):
            d = delivered.get(key, {}).get(m, 0.0)
            now = db_prev.get(key, {}).get(m, 0.0)
            if abs(now - d) > RESTATE_TOLERANCE:
                rows.append({"product": key[0], "currency": key[1], "stage": key[2], "measure": m,
                             "delivered": d, "database_now": now, "delta": now - d})
    return {"previous_file": path.name, "rows": rows}


# ---------------------------------------------------------------- Excel
def make_workbook(path: Path, report_date: dt.date, curr: dict, comparison: list[dict]) -> None:
    wb = Workbook()
    bold = Font(bold=True)
    ws = wb.active
    ws.title = "Basis"
    ws.append(["value_date", "product", "currency", "stage", "exposure", "ecl"])
    for key in sorted(curr):
        ws.append([report_date.isoformat(), *key, curr[key]["exposure"], curr[key]["ecl"]])
    ws.append([None, "TOTAL", None, None, f"=SUM(E2:E{ws.max_row})", f"=SUM(F2:F{ws.max_row})"])
    for cell in ws[1] + ws[ws.max_row]:
        cell.font = bold

    wc = wb.create_sheet("Comparison")
    headers = ["product", "currency", "stage", "status"] + [f"{m}_{s}" for m in ("exposure", "ecl")
                                                            for s in ("prev", "curr", "delta", "pct")]
    wc.append(headers)
    amber, red = PatternFill("solid", fgColor="FFE699"), PatternFill("solid", fgColor="F4B084")
    for r in comparison:
        wc.append([r[h] for h in headers])
        for col in ("exposure_pct", "ecl_pct"):
            value = r[col]
            if value is not None and abs(value) >= 0.05:
                wc.cell(wc.max_row, headers.index(col) + 1).fill = red if abs(value) >= 0.10 else amber
    for cell in wc[1]:
        cell.font = bold

    wb.create_sheet("Query").append([EXTRACT_SQL])
    for sheet in wb.worksheets:
        for col in sheet.columns:
            sheet.column_dimensions[col[0].column_letter].width = 18
    wb.save(path)


def save_verified_workbook(target: Path, report_date: dt.date, curr: dict, comparison: list[dict]) -> None:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_file = Path(tmp) / target.name
        make_workbook(tmp_file, report_date, curr, comparison)
        if read_delivered_file(tmp_file) != {k: {m: float(v[m]) for m in v} for k, v in curr.items()}:
            raise BasisError("The Excel file does not reconcile with the database – nothing was saved.")
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(tmp_file, target)


# ---------------------------------------------------------------- notification
def build_message(summary: dict, comment: str | None, test: bool) -> str:
    status = summary["status"]
    date = summary["report_date"]
    prefix = "[TEST] " if test else ""
    readiness = summary.get("readiness", {})
    if status == "not_ready":
        return (f"{prefix}⚠ Provision basis {date} – data not ready (attempt {summary.get('attempt', 1)}/3)\n\n"
                f"- Table rebuilt after month-end: {readiness.get('rebuilt_after_month_end')}\n"
                f"- Rows for {date}: {readiness.get('row_count')}\n")
    if status == "error":
        return f"{prefix}❌ Provision basis {date} failed\n\n{summary.get('error')}\n"

    missing = readiness.get("missing_products", [])
    title = f"{prefix}✔ Provision basis {date} saved"
    if missing:
        title += f" – {', '.join(missing)} missing"
    lines = [title, "", "**Checks**",
             f"- Table rebuilt {readiness['rebuilt_at']} ✔",
             f"- {readiness['row_count']} rows for {date} ✔"]
    for d in readiness.get("missing_details", []):
        lines.append(f"- ⚠ {d['product']} missing (last month {format_millions(d['prev_exposure'], d['currency'])} exposure, "
                     f"{format_millions(d['prev_ecl'], d['currency'])} ECL, last rows {d['last_value_date']})")
    for p in readiness.get("new_products", []):
        lines.append(f"- ℹ New product: {p}")
    restated = summary["restatements"]["rows"]
    lines.append(f"- {'⚠' if restated else '✔'} {len(restated)} restated row(s) vs "
                 f"{summary['restatements']['previous_file'] or 'no previous file'}")

    lines += ["", "**Comment**"]
    if comment:
        lines.append(comment.strip())
    else:
        for n in summary["notable_changes"][:6]:
            lines.append(f"- {n['product']} stage {n['stage']} {n['measure']}: {format_millions(n['delta'], n['currency'])} "
                         f"({format_pct(n['pct'])}) {n['status']}".rstrip())

    lines += ["", "**Exposure and ECL per product and stage**", "",
              "| Product | Stage | Exposure | Δ % | ECL | Δ % |", "|---|---|---:|---:|---:|---:|"]
    for t in summary["totals_by_product"]:
        lines.append(f"| **{t['product']}** | **all** | **{format_millions(t['exposure_curr'], t['currency'])}** | "
                     f"**{format_pct(t['exposure_pct'])}** | **{format_millions(t['ecl_curr'], t['currency'])}** | "
                     f"**{format_pct(t['ecl_pct'])}** |")
        for r in (r for r in summary["comparison"] if r["product"] == t["product"]):
            lines.append(f"| | {r['stage']} | {format_millions(r['exposure_curr'], r['currency'])} | {format_pct(r['exposure_pct'])} "
                         f"| {format_millions(r['ecl_curr'], r['currency'])} | {format_pct(r['ecl_pct'])} |")
    return "\n".join(lines) + "\n"


def send_or_save(message: str, out_file: Path, dry_run: bool) -> str:
    out_file.write_text(message, encoding="utf-8")
    url = os.environ.get("PROVISION_WEBHOOK_URL")
    if dry_run or not url:
        return "saved"
    try:
        req = urllib.request.Request(url, data=json.dumps({"text": message}).encode(),
                                     headers={"Content-Type": "application/json"})
        urllib.request.urlopen(req, timeout=30).close()
        return "sent"
    except Exception as exc:  # noqa: BLE001 – a failed notification must never fail the build
        print(f"Notification failed: {exc}", file=sys.stderr)
        return "failed"


# ---------------------------------------------------------------- commands
def run_build(args) -> int:
    today = args.today or dt.date.today()
    report_date = args.date or previous_month_end(today)
    prev_date = previous_month_end(report_date)
    temp_root = Path(tempfile.gettempdir()) / "provision-basis"
    output_dir = temp_root / "output" if args.dry_run else args.output_dir
    target = output_dir / basis_filename(report_date)
    run_dir = (temp_root / "runs" if args.dry_run else RUNS_DIR) / report_date.isoformat()
    run_dir.mkdir(parents=True, exist_ok=True)
    summary_path = run_dir / "summary.json"
    summary = {"report_date": report_date.isoformat(), "previous_date": prev_date.isoformat(),
               "mode": "manual" if args.manual else "scheduled"}

    if not args.manual:
        bday = business_day_of_month(today)
        if bday not in BUSINESS_DAYS:
            print(f"Business day {bday} – nothing to do.\nSTATUS=skipped")
            return 3
    if target.exists() and not args.overwrite:
        print(f"{target.name} already exists.\nSTATUS=skipped")
        return 3

    try:
        with connect(args.db) as conn:
            curr, prev = load_totals(conn, report_date), load_totals(conn, prev_date)
            readiness = data_checks(conn, report_date, curr, prev)
        summary["readiness"] = readiness
        if not readiness["ready"] and not args.force:
            summary["status"] = "not_ready"
            summary_path.write_text(json.dumps(round_floats(summary), indent=2), encoding="utf-8")
            if not args.manual:
                print(f"NOTIFY={send_or_save(build_message(summary, None, False), run_dir / 'notification.md', args.dry_run)}")
            print(f"STATUS=not_ready\nSUMMARY_JSON={summary_path}")
            return 2

        comparison = compare_months(curr, prev)
        save_verified_workbook(target, report_date, curr, comparison)
        summary |= {
            "status": "done",
            "report_file": target.name,
            "totals_by_product": sum_by(comparison, ("product", "currency")),
            "comparison": comparison,
            "notable_changes": flag_changes(comparison),
            "restatements": find_restatements(output_dir, prev_date, prev),
            "warnings": ["Saved with --force although data checks failed"] if not readiness["ready"] else [],
        }
        summary_path.write_text(json.dumps(round_floats(summary), indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"STATUS=done\nREPORT_FILE={target}\nSUMMARY_JSON={summary_path}")
        return 0
    except (BasisError, sqlite3.Error, OSError) as exc:
        summary |= {"status": "error", "error": str(exc)}
        summary_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")
        if not args.manual:
            print(f"NOTIFY={send_or_save(build_message(summary, None, False), run_dir / 'notification.md', args.dry_run)}")
        print(f"ERROR: {exc}\nSTATUS=error", file=sys.stderr)
        return 1


def run_notify(args) -> int:
    summary = json.loads(Path(args.summary).read_text(encoding="utf-8"))
    comment = Path(args.comment_file).read_text(encoding="utf-8") if args.comment_file else None
    message = build_message(summary, comment, args.test)
    result = send_or_save(message, Path(args.summary).parent / "notification.md", args.dry_run)
    print(message)
    print(f"NOTIFY={result}")
    return 0


def main(argv: list[str] | None = None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)
    b = sub.add_parser("build")
    b.add_argument("--manual", action="store_true", help="Run now, ignoring the business-day rule")
    b.add_argument("--date", type=dt.date.fromisoformat, help="Report date (a month-end). Default: last month-end")
    b.add_argument("--force", action="store_true", help="Save even if the data checks fail")
    b.add_argument("--overwrite", action="store_true", help="Replace an existing file for the same month")
    b.add_argument("--dry-run", action="store_true", help="Save to a temp folder and do not send anything")
    b.add_argument("--db", type=Path, default=Path(os.environ.get("PROVISION_DB", DEFAULT_DB)))
    b.add_argument("--output-dir", type=Path, default=Path(os.environ.get("PROVISION_OUTPUT_DIR", DEFAULT_OUTPUT)))
    b.add_argument("--today", type=dt.date.fromisoformat, help=argparse.SUPPRESS)
    n = sub.add_parser("notify")
    n.add_argument("--summary", required=True)
    n.add_argument("--comment-file")
    mode = n.add_mutually_exclusive_group()
    mode.add_argument("--dry-run", action="store_true", help="Only save notification.md")
    mode.add_argument("--test", action="store_true", help="Send with [TEST] in the title")
    args = parser.parse_args(argv)
    return run_build(args) if args.command == "build" else run_notify(args)


if __name__ == "__main__":
    sys.exit(main())
