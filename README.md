# AI-Based Internship Review Automation System

## Workflow
Excel student list → reminder email → student submits GitHub/recording links → AI-assisted repository review → feedback email and next review date → repeat until completion.

## Setup
1. Install Python 3.10 or newer.
2. Extract this ZIP and open a terminal in the extracted folder.
3. Install packages: `pip install -r requirements.txt`
4. Configure environment variables listed in `.env.example`. Do not commit real credentials.
5. Create the working Excel file: `python internship_review.py init`
6. Replace the sample student row with real details and set review dates as `YYYY-MM-DD`.

## Commands
- `python internship_review.py init` — creates `students.xlsx`.
- `python internship_review.py remind` — emails students whose review date is today.
- `python internship_review.py review` — reviews rows marked `Awaiting Submission` that have a GitHub link, emails feedback, and schedules the next review.
- `python internship_review.py all` — runs reminders, then reviews.

## Excel workflow note
The included `students_template.xlsx` is a blank column template. The script uses `students.xlsx`; create it with the `init` command or rename/copy the template to `students.xlsx`. After a student submits the form, copy their GitHub URL into the matching row's `GitHub Link` cell before running `review`.

## Requirements and limitations
The script summarizes public GitHub repository metadata, recent commits, file paths, and README text; it does not download or execute the student's code. Automatic Google/Microsoft Forms response import is not implemented. Private repositories may require a token with appropriate permissions. Test with a test student and test email before production use.
