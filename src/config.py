"""Shared constants so the data generator, tracker builder and report agree."""

STAGES = ["Applied", "Screened", "Interviewed", "Offered", "Joined"]
OUTCOMES = ["In Progress", "Rejected", "Withdrawn", "Joined"]
ROLES = [
    "Data Analyst Intern",
    "Software Engineering Intern",
    "Marketing Intern",
    "Operations Intern",
    "Finance Intern",
]
SOURCES = ["LinkedIn", "Job Portal", "Employee Referral", "Campus Drive", "Company Website"]
RECRUITERS = ["Recruiter A", "Recruiter B", "Recruiter C"]

CSV_COLUMNS = [
    "candidate_id", "name", "email", "role", "source", "recruiter",
    "applied_date", "screening_date", "interview_date", "offer_date",
    "stage", "outcome", "notes",
]

# Target days used to flag slow steps in the tracker and the report.
SCREEN_SLA_DAYS = 5
