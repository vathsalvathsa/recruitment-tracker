"""Generate a fully synthetic candidate dataset (no real people).

Names are drawn from short fictional lists and emails use the reserved
example.com domain, so the sample data is safe to publish.

Usage:
    python -m src.generate_sample_data --rows 120 --seed 42
"""
import argparse
import csv
import random
from datetime import date, timedelta
from pathlib import Path

from .config import CSV_COLUMNS, RECRUITERS, ROLES, SOURCES, STAGES

FIRST = ["Aarav", "Diya", "Kabir", "Meera", "Rohan", "Isha", "Vivaan", "Anaya", "Arjun", "Sara",
         "Neel", "Tara", "Yash", "Riya", "Dev", "Kiara", "Ishaan", "Mira", "Aditya", "Zoya"]
LAST = ["Sharma", "Verma", "Gupta", "Mehta", "Singh", "Iyer", "Nair", "Khanna", "Bose", "Rao",
        "Malik", "Joshi", "Kapoor", "Reddy", "Chopra"]

# Chance that a candidate's furthest stage is Applied, Screened, Interviewed, Offered, Joined.
STAGE_WEIGHTS = [0.30, 0.30, 0.22, 0.10, 0.08]


def build_rows(n: int, seed: int) -> list[dict]:
    rng = random.Random(seed)
    start = date(2026, 6, 1)
    rows = []
    for i in range(1, n + 1):
        name = f"{rng.choice(FIRST)} {rng.choice(LAST)}"
        stage_idx = rng.choices(range(len(STAGES)), weights=STAGE_WEIGHTS)[0]
        applied = start + timedelta(days=rng.randint(0, 75))
        screened = interview = offer = None
        if stage_idx >= 1:
            screened = applied + timedelta(days=rng.randint(1, 8))
        if stage_idx >= 2:
            interview = screened + timedelta(days=rng.randint(2, 9))
        if stage_idx >= 3:
            offer = interview + timedelta(days=rng.randint(2, 7))

        if stage_idx == 4:
            outcome = "Joined"
        elif stage_idx == 3:
            outcome = rng.choices(["In Progress", "Rejected", "Withdrawn"], [0.6, 0.1, 0.3])[0]
        else:
            outcome = rng.choices(["In Progress", "Rejected", "Withdrawn"], [0.35, 0.55, 0.10])[0]

        rows.append({
            "candidate_id": f"C{i:04d}",
            "name": name,
            "email": f"{name.lower().replace(' ', '.')}.{i}@example.com",
            "role": rng.choice(ROLES),
            "source": rng.choices(SOURCES, weights=[0.30, 0.25, 0.15, 0.15, 0.15])[0],
            "recruiter": rng.choice(RECRUITERS),
            "applied_date": applied.isoformat(),
            "screening_date": screened.isoformat() if screened else "",
            "interview_date": interview.isoformat() if interview else "",
            "offer_date": offer.isoformat() if offer else "",
            "stage": STAGES[stage_idx],
            "outcome": outcome,
            "notes": "",
        })
    return rows


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--rows", type=int, default=120)
    p.add_argument("--seed", type=int, default=42)
    p.add_argument("--out", default="data/sample_candidates.csv")
    args = p.parse_args()

    rows = build_rows(args.rows, args.seed)
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
        w.writeheader()
        w.writerows(rows)
    print(f"Wrote {len(rows)} synthetic candidates to {out}")


if __name__ == "__main__":
    main()
