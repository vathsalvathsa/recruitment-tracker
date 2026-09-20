"""Build a recruitment pipeline report (Markdown + funnel chart) from the candidate CSV.

Usage:
    python -m src.weekly_report --input data/sample_candidates.csv --outdir output
"""
import argparse
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

from .config import SCREEN_SLA_DAYS, STAGES

DATE_COLS = ["applied_date", "screening_date", "interview_date", "offer_date"]


def load(path: str) -> pd.DataFrame:
    df = pd.read_csv(path, parse_dates=DATE_COLS)
    validate(df)
    df["stage_order"] = df["stage"].map({s: i + 1 for i, s in enumerate(STAGES)})
    return df


def validate(df: pd.DataFrame) -> None:
    """Data-quality checks: recruitment records must be accurate."""
    if df["candidate_id"].duplicated().any():
        raise ValueError("Duplicate candidate_id values found")
    bad_stage = set(df["stage"]) - set(STAGES)
    if bad_stage:
        raise ValueError(f"Unknown stage values: {sorted(bad_stage)}")
    for earlier, later in zip(DATE_COLS, DATE_COLS[1:]):
        wrong = df[df[later].notna() & df[earlier].notna() & (df[later] < df[earlier])]
        if not wrong.empty:
            raise ValueError(f"{later} is before {earlier} for {list(wrong['candidate_id'])}")


def funnel(df: pd.DataFrame) -> pd.DataFrame:
    """Candidates who reached at least each stage, with step and overall conversion."""
    counts = [int((df["stage_order"] >= i).sum()) for i in range(1, len(STAGES) + 1)]
    out = pd.DataFrame({"stage": STAGES, "reached": counts})
    out["step_conversion"] = (out["reached"] / out["reached"].shift(1)).round(3)
    out["overall_conversion"] = (out["reached"] / out["reached"].iloc[0]).round(3)
    return out


def by_source(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("source")
    out = pd.DataFrame({
        "applicants": g.size(),
        "interviewed_plus": g["stage_order"].apply(lambda s: int((s >= 3).sum())),
        "offered_plus": g["stage_order"].apply(lambda s: int((s >= 4).sum())),
        "joined": g["outcome"].apply(lambda s: int((s == "Joined").sum())),
    })
    out["interview_rate"] = (out["interviewed_plus"] / out["applicants"]).round(3)
    out["offer_rate"] = (out["offered_plus"] / out["applicants"]).round(3)
    return out.sort_values("applicants", ascending=False).reset_index()


def timing(df: pd.DataFrame) -> dict:
    to_screen = (df["screening_date"] - df["applied_date"]).dt.days.dropna()
    to_offer = (df["offer_date"] - df["applied_date"]).dt.days.dropna()
    return {
        "avg_days_to_screen": round(float(to_screen.mean()), 1),
        "avg_days_to_offer": round(float(to_offer.mean()), 1),
        "pct_screened_within_sla": round(float((to_screen <= SCREEN_SLA_DAYS).mean()), 3),
    }


def to_markdown_table(df: pd.DataFrame) -> str:
    cols = list(df.columns)
    lines = ["| " + " | ".join(cols) + " |", "|" + "|".join(["---"] * len(cols)) + "|"]
    for _, r in df.iterrows():
        lines.append("| " + " | ".join("-" if pd.isna(r[c]) else str(r[c]) for c in cols) + " |")
    return "\n".join(lines)


def save_chart(f: pd.DataFrame, path: Path) -> None:
    fig, ax = plt.subplots(figsize=(7, 4))
    bars = ax.bar(f["stage"], f["reached"], color="#0A7EA4")
    ax.bar_label(bars, padding=2)
    ax.set_title("Hiring funnel: candidates reaching each stage")
    ax.set_ylabel("Candidates")
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


def build_report(df: pd.DataFrame, outdir: Path) -> Path:
    outdir.mkdir(parents=True, exist_ok=True)
    f, s, t = funnel(df), by_source(df), timing(df)
    save_chart(f, outdir / "funnel.png")
    outcomes = df["outcome"].value_counts().rename_axis("outcome").reset_index(name="candidates")
    md = f"""# Recruitment Pipeline Report

Total candidates: **{len(df)}** | Data period: {df['applied_date'].min():%d %b %Y} to {df['applied_date'].max():%d %b %Y}

## Hiring funnel
{to_markdown_table(f)}

![Funnel](funnel.png)

## Sourcing channel effectiveness
{to_markdown_table(s)}

## Speed of process
- Average days from application to screening: **{t['avg_days_to_screen']}**
- Average days from application to offer: **{t['avg_days_to_offer']}**
- Screened within {SCREEN_SLA_DAYS} days: **{t['pct_screened_within_sla']:.0%}**

## Current outcomes
{to_markdown_table(outcomes)}

_All data in this report is synthetic._
"""
    path = outdir / "weekly_report.md"
    path.write_text(md, encoding="utf-8")
    return path


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--input", default="data/sample_candidates.csv")
    p.add_argument("--outdir", default="output")
    a = p.parse_args()
    print(f"Report written to {build_report(load(a.input), Path(a.outdir))}")


if __name__ == "__main__":
    main()
