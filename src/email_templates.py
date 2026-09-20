"""Fill a candidate-communication template from a row of the candidate CSV.

Usage:
    python -m src.email_templates --id C0001 --template interview_invite \
        --slot "Mon 10 Aug 2026, 11:00 AM" --mode "In person, Gurugram office"
"""
import argparse
import csv
from pathlib import Path
from string import Template

TEMPLATE_DIR = Path(__file__).resolve().parent.parent / "templates"


def available_templates() -> list[str]:
    return sorted(p.stem for p in TEMPLATE_DIR.glob("*.txt"))


def render(template_name: str, candidate: dict, **extra: str) -> str:
    path = TEMPLATE_DIR / f"{template_name}.txt"
    if not path.exists():
        raise FileNotFoundError(f"Unknown template '{template_name}'. Available: {available_templates()}")
    values = {**candidate, "first_name": candidate["name"].split()[0], **extra}
    # safe_substitute leaves unfilled placeholders visible so nothing is silently dropped
    return Template(path.read_text(encoding="utf-8")).safe_substitute(values)


def find_candidate(csv_path: str, candidate_id: str) -> dict:
    with open(csv_path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if row["candidate_id"] == candidate_id:
                return row
    raise KeyError(f"Candidate {candidate_id} not found")


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--input", default="data/sample_candidates.csv")
    p.add_argument("--id", required=True)
    p.add_argument("--template", required=True, choices=available_templates())
    p.add_argument("--slot", default="$slot")
    p.add_argument("--mode", default="$mode")
    p.add_argument("--deadline", default="$deadline")
    a = p.parse_args()
    cand = find_candidate(a.input, a.id)
    print(render(a.template, cand, slot=a.slot, mode=a.mode, deadline=a.deadline))


if __name__ == "__main__":
    main()
