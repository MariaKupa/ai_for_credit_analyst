"""Generate a fully synthetic IFRS 9 portfolio for the month-end provision basis demo.

Every product, amount and event here is invented. The data deliberately contains the
situations the automation has to handle:

* a stage 2 migration in Product B (a "notable change")
* Product C stops reporting before the latest month-end ("missing product")
* optionally a late restatement of the previous month-end (--apply-restatement)
"""
from __future__ import annotations

import argparse
import datetime as dt
import random
import sqlite3
from pathlib import Path

DEFAULT_DB = Path(__file__).resolve().parent / "portfolio.db"

MONTH_ENDS = [dt.date(2026, 6, 30), dt.date(2026, 7, 31), dt.date(2026, 8, 31),
              dt.date(2026, 9, 30), dt.date(2026, 10, 31)]

# product: (currency, exposure at first month-end, stage mix, first month-end, last month-end)
PRODUCTS = {
    "Product A": ("EUR", 1_650_000_000, (0.88, 0.09, 0.03), None, None),
    "Product B": ("USD",   310_000_000, (0.90, 0.08, 0.02), None, None),
    "Product C": ("EUR",    12_000_000, (0.95, 0.04, 0.01), None, dt.date(2026, 9, 30)),
}

# Typical ECL coverage ranges per stage (12-month ECL in stage 1, lifetime in 2 and 3).
COVERAGE = {1: (0.004, 0.012), 2: (0.04, 0.08), 3: (0.40, 0.60)}


def build_rows(seed: int) -> list[tuple]:
    rng = random.Random(seed)
    rows = []
    for product, (ccy, start_exposure, mix, first, last) in PRODUCTS.items():
        exposure = start_exposure
        coverage = {s: rng.uniform(*COVERAGE[s]) for s in COVERAGE}
        stage_mix = list(mix)
        for month_end in MONTH_ENDS:
            if (first and month_end < first) or (last and month_end > last):
                continue
            exposure *= 1 + rng.uniform(-0.02, 0.03)
            if product == "Product B" and month_end == MONTH_ENDS[-1]:
                stage_mix = [stage_mix[0] - 0.03, stage_mix[1] + 0.03, stage_mix[2]]  # stage 2 migration
            for stage, share in zip((1, 2, 3), stage_mix):
                if share <= 0:
                    continue
                stage_exposure = exposure * share
                cov = coverage[stage] * (1 + rng.uniform(-0.03, 0.03))
                rows.append((month_end.isoformat(), product, ccy, stage,
                             round(stage_exposure, 2), round(stage_exposure * cov, 2)))
    return rows


def create(db: Path, seed: int, rebuilt_at: str) -> None:
    db.unlink(missing_ok=True)
    with sqlite3.connect(db) as conn:
        conn.executescript("""
            CREATE TABLE data_source (
                value_date TEXT NOT NULL, product TEXT NOT NULL, currency TEXT NOT NULL,
                stage INTEGER NOT NULL, exposure REAL NOT NULL, ecl REAL NOT NULL
            );
            CREATE TABLE load_log (table_name TEXT PRIMARY KEY, rebuilt_at TEXT NOT NULL);
        """)
        conn.executemany("INSERT INTO data_source VALUES (?,?,?,?,?,?)", build_rows(seed))
        conn.execute("INSERT INTO load_log VALUES ('data_source', ?)", (rebuilt_at,))
    print(f"Created {db}")


def apply_restatement(db: Path, rebuilt_at: str) -> None:
    """Simulate a late correction of last month's numbers after the file was delivered."""
    with sqlite3.connect(db) as conn:
        cur = conn.execute("""UPDATE data_source SET ecl = ecl * 1.03
                              WHERE value_date = ? AND product = 'Product A'
                                AND stage = 2""", (MONTH_ENDS[-2].isoformat(),))
        conn.execute("UPDATE load_log SET rebuilt_at = ? WHERE table_name = 'data_source'", (rebuilt_at,))
    print(f"Restated {cur.rowcount} row(s) for {MONTH_ENDS[-2]}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--db", type=Path, default=DEFAULT_DB)
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--rebuilt-at", default=dt.datetime.now().isoformat(timespec="seconds"),
                        help="Timestamp of the last table rebuild (used by the readiness check)")
    parser.add_argument("--apply-restatement", action="store_true",
                        help="Change last month's ECL in an existing database instead of creating it")
    args = parser.parse_args()
    if args.apply_restatement:
        apply_restatement(args.db, args.rebuilt_at)
    else:
        create(args.db, args.seed, args.rebuilt_at)


if __name__ == "__main__":
    main()
