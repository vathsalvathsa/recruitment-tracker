# Recruitment Tracker & Hiring Analytics Toolkit

A small Python and Excel toolkit that covers the everyday work of a Talent Acquisition team: keeping an accurate **candidate tracker**, producing **pipeline reports**, and sending consistent **candidate emails**.

I built it as a personal portfolio project to combine my data analysis background (Python, pandas, SQL, Power BI, Tableau) with the workflow of a recruitment coordinator. **All data is synthetic** (fictional names, `example.com` emails). No real candidate data is included.

## What it does

| Need in recruitment | What this project provides |
|---|---|
| Maintain recruitment trackers | Excel workbook with a candidate log, dropdown lists, date validation and colour-coded outcomes |
| Maintain reports | Dashboard sheet where every figure is a live formula (funnel, source effectiveness, outcomes by role, time to screen) |
| Accuracy in recruitment documentation | Input validation in Excel, plus a Python data-quality check for duplicate IDs, unknown stages and out-of-order dates |
| Source and screen candidates | Sourcing-channel analysis showing which channels reach interview and offer stages |
| Candidate communication | Email templates (application received, interview invite, rejection, onboarding documents) filled from tracker data |
| Time management | Screening SLA setting that highlights candidates waiting too long |

## Project structure

```
recruitment-tracker/
├── src/
│   ├── config.py                 # stages, roles, sources, SLA
│   ├── generate_sample_data.py   # synthetic candidate data
│   ├── build_tracker.py          # Excel tracker + dashboard (openpyxl)
│   ├── weekly_report.py          # pandas report + funnel chart
│   └── email_templates.py        # fill candidate emails from a template
├── templates/                    # plain-text email templates
├── data/sample_candidates.csv    # 120 synthetic candidates
├── output/                       # example generated outputs
└── tests/test_report.py          # unit tests
```

## Quick start

```bash
pip install -r requirements.txt

python -m src.generate_sample_data --rows 120 --seed 42
python -m src.build_tracker          # output/recruitment_tracker.xlsx
python -m src.weekly_report          # output/weekly_report.md + funnel.png
python -m src.email_templates --id C0001 --template interview_invite \
    --slot "Mon 10 Aug 2026, 11:00 AM" --mode "In person"

python -m unittest discover -s tests -t .
```

To use your own data, replace `data/sample_candidates.csv` with a file that has the same columns (see `src/config.py`).

## Example output

The report on the sample data shows a 120 → 78 → 47 → 17 → 6 funnel (Applied → Screened → Interviewed → Offered → Joined). See `output/weekly_report.md` and `output/funnel.png`.

## Design notes

- **Stage vs Outcome:** *Stage* is the furthest step reached; *Outcome* is the current status. This lets the funnel count everyone who reached a stage, even if they were later rejected.
- **Formulas, not pasted values:** the workbook recalculates when a recruiter edits a row. Blue text marks typed inputs; black text marks formulas.
- **Excel and pandas agree:** the dashboard figures match the Python report (checked while building).

## Skills demonstrated

Python, pandas, openpyxl, Excel formulas (COUNTIF/COUNTIFS/AVERAGE/MATCH), data validation and conditional formatting, data quality checks, matplotlib, unit testing, documentation.

## Licence

MIT. See `LICENSE`.
