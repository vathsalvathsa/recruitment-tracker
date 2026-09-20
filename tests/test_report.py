import unittest
from datetime import date

import pandas as pd

from src.config import STAGES
from src.generate_sample_data import build_rows
from src.weekly_report import by_source, funnel, timing, validate


def make_df(n=200, seed=1):
    df = pd.DataFrame(build_rows(n, seed))
    for c in ["applied_date", "screening_date", "interview_date", "offer_date"]:
        df[c] = pd.to_datetime(df[c].replace("", None))
    df["stage_order"] = df["stage"].map({s: i + 1 for i, s in enumerate(STAGES)})
    return df


class ReportTests(unittest.TestCase):
    def test_generated_data_passes_validation(self):
        validate(make_df())

    def test_funnel_is_monotonic_and_starts_with_everyone(self):
        df = make_df()
        f = funnel(df)
        self.assertEqual(f["reached"].iloc[0], len(df))
        self.assertTrue(f["reached"].is_monotonic_decreasing)

    def test_source_applicants_sum_to_total(self):
        df = make_df()
        self.assertEqual(by_source(df)["applicants"].sum(), len(df))

    def test_timing_values_are_positive(self):
        t = timing(make_df())
        self.assertGreater(t["avg_days_to_screen"], 0)
        self.assertGreater(t["avg_days_to_offer"], t["avg_days_to_screen"])

    def test_validation_catches_bad_dates(self):
        df = make_df()
        i = df.index[df["screening_date"].notna()][0]
        df.loc[i, "screening_date"] = df.loc[i, "applied_date"] - pd.Timedelta(days=3)
        with self.assertRaises(ValueError):
            validate(df)

    def test_validation_catches_duplicate_ids(self):
        df = make_df()
        df.loc[1, "candidate_id"] = df.loc[0, "candidate_id"]
        with self.assertRaises(ValueError):
            validate(df)


if __name__ == "__main__":
    unittest.main()
