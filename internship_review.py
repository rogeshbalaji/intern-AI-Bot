"""AI-Based Internship Review Automation System.

Commands:
  python internship_review.py init
  python internship_review.py remind
  python internship_review.py review
  python internship_review.py all

Configure ANTHROPIC_API_KEY, SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS,
FORM_URL, and optionally GITHUB_TOKEN / CLAUDE_MODEL as environment variables.

Note: Form responses are not imported automatically. Copy the student's GitHub
URL from the form into the matching Excel row before running `review`.
"""
import os
import re
import smtplib
import ssl
import sys
from datetime import date, datetime, timedelta
from email.message import EmailMessage

import anthropic
import pandas as pd
import requests

EXCEL_FILE = "students.xlsx"
REVIEW_GAP_DAYS = 7
TOTAL_REVIEWS = 4
REVIEW_TIME = "10:00"
MODEL = os.getenv("CLAUDE_MODEL", "claude-sonnet-4-5")
FORM_URL = os.getenv("FORM_URL", "https://forms.gle/your-form-link")
COLUMNS = [
    "Name", "Email", "Review Date", "Review Time", "Review No", "Status",
    "GitHub Link", "Recording Link", "Feedback",
]
REVIEW_CRITERIA = """
1. Progress since the last review (commit history, new features)
2. Code quality and structure (readability, organisation, naming)
3. README and documentation (setup steps, project description)
4. Completeness against a typical internship project scope
5. Best practices (tests, error handling, no secrets committed)
"""


def load_df():
    if not os.path.exists(EXCEL_FILE):
        raise FileNotFoundError(f"{EXCEL_FILE} not found. Run: python internship_review.py init")
    df = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df


def save_df(df):
    df.to_excel(EXCEL_FILE, index=False)


def cmd_init():
    if os.path.exists(EXCEL_FILE):
        print(f"{EXCEL_FILE} already exists - not overwriting.")
        return
    sample = {
        "Name": "Student One",
        "Email": "student@example.com",
        "Review Date": date.today().isoformat(),
        "Review Time": REVIEW_TIME,
        "Review No": "1",
        "Status": "Scheduled",
    }
    pd.DataFrame([sample], columns=COLUMNS).fillna("").to_excel(EXCEL_FILE, index=False)
    print(f"Created {EXCEL_FILE}. Replace the sample student details before use.")


def send_email(to, subject, body):
    required = ["SMTP_HOST", "SMTP_USER", "SMTP_PASS"]
    missing = [key for key in required if not os.getenv(key)]
    if missing:
        raise RuntimeError("Missing email configuration: " + ", ".join(missing))
    msg = EmailMessage()
    msg["From"] = os.environ["SMTP_USER"]
    msg["To"] = to
    msg["Subject"] = subject
    msg.set_content(body)
    context = ssl.create_default_context()
    port = int(os.getenv("SMTP_PORT", "465"))
    with smtplib.SMTP_SSL(os.environ["SMTP_HOST"], port, context=context) as server:
        server.login(os.environ["SMTP_USER"], os.environ["SMTP_PASS"])
        server.send_message(msg)


def cmd_remind():
    df = load_df()
    today = date.today().isoformat()
    for i, row in df.iterrows():
        if row["Status"] == "Scheduled" and str(row["Review Date"])[:10] == today:
            body = (
                f"Hi {row['Name']},\n\n"
                f"Your internship review #{row['Review No']} is scheduled for "
                f"{str(row['Review Date'])[:10]} at {row['Review Time']}.\n\n"
                f"Please submit your GitHub repository and screen recording links here:\n"
                f"{FORM_URL}\n\nRegards,\nInternship Coordinator"
            )
            send_email(row["Email"], f"Internship Review #{row['Review No']} - Submission Link", body)
            df.at[i, "Status"] = "Awaiting Submission"
            print(f"Reminder sent to {row['Name']}")
    save_df(df)


def github_summary(url):
    match = re.search(r"github\.com/([^/\\s]+)/([^/\\s#?]+)", url)
    if not match:
        return f"Could not parse GitHub link: {url}"
    owner, repo = match.group(1), match.group(2).removesuffix(".git")
    headers = {"Accept": "application/vnd.github+json"}
    token = os.getenv("GITHUB_TOKEN")
    if token:
        headers["Authorization"] = f"Bearer {token}"
    base = f"https://api.github.com/repos/{owner}/{repo}"

    def get(path, **kwargs):
        request_headers = {**headers, **kwargs.pop("headers", {})}
        return requests.get(base + path, headers=request_headers, timeout=30, **kwargs)

    response = get("")
    if not response.ok:
        return f"Could not access repository (HTTP {response.status_code}). Check the URL and repository access."
    info = response.json()
    parts = [
        f"Repository: {info.get('full_name')} | language: {info.get('language')} | "
        f"last push: {info.get('pushed_at')}"
    ]
    response = get("/commits", params={"per_page": 15})
    if response.ok:
        commits = [
            f"- {item['commit']['author']['date'][:10]}: "
            f"{item['commit']['message'].splitlines()[0]}"
            for item in response.json()
        ]
        parts.append("Recent commits:\n" + "\n".join(commits))
    response = get("/git/trees/HEAD", params={"recursive": "1"})
    if response.ok:
        paths = [item["path"] for item in response.json().get("tree", []) if item["type"] == "blob"][:80]
        parts.append("Files:\n" + "\n".join(paths))
    response = get("/readme", headers={"Accept": "application/vnd.github.raw+json"})
    parts.append("README (first 3000 chars):\n" + (response.text[:3000] if response.ok else "No README found."))
    return "\n\n".join(parts)


def ai_review(name, review_no, repo_info, previous_feedback):
    if not os.getenv("ANTHROPIC_API_KEY"):
        raise RuntimeError("Missing ANTHROPIC_API_KEY environment variable.")
    prompt = f"""You are reviewing an intern's project for a coordinator.
Student: {name}
Review number: {review_no} of {TOTAL_REVIEWS}

Review criteria:
{REVIEW_CRITERIA}

Previous feedback (check whether it was addressed):
{previous_feedback or 'None - this is the first review.'}

Repository data:
{repo_info}

Write concise, constructive plain-text email feedback with sections:
SUMMARY (2-3 lines), STRENGTHS, CHANGES REQUIRED (numbered and actionable),
SCORE (out of 10). Only comment on evidence shown; note anything that cannot
be judged from the available data."""
    client = anthropic.Anthropic()
    response = client.messages.create(
        model=MODEL,
        max_tokens=1200,
        messages=[{"role": "user", "content": prompt}],
    )
    return response.content[0].text


def cmd_review():
    df = load_df()
    for i, row in df.iterrows():
        if row["Status"] != "Awaiting Submission" or not row["GitHub Link"].strip():
            continue
        print(f"Reviewing {row['Name']}...")
        repo_info = github_summary(row["GitHub Link"].strip())
        feedback = ai_review(row["Name"], row["Review No"], repo_info, row["Feedback"])
        try:
            review_no = int(row["Review No"] or 1)
        except ValueError:
            print(f"Skipping {row['Name']}: Review No must be numeric.")
            continue

        if review_no >= TOTAL_REVIEWS:
            footer = "\n\nThis was your final review. Congratulations on completing the internship!"
            df.at[i, "Status"] = "Completed"
        else:
            next_date = (datetime.now() + timedelta(days=REVIEW_GAP_DAYS)).date().isoformat()
            footer = (
                f"\n\nNEXT REVIEW: #{review_no + 1} on {next_date} at {REVIEW_TIME}\n"
                f"Submission link: {FORM_URL}"
            )
            df.at[i, "Review No"] = str(review_no + 1)
            df.at[i, "Review Date"] = next_date
            df.at[i, "Review Time"] = REVIEW_TIME
            df.at[i, "Status"] = "Scheduled"

        df.at[i, "Feedback"] = feedback
        df.at[i, "GitHub Link"] = ""
        send_email(
            row["Email"],
            f"Internship Review #{review_no} - Feedback",
            f"Hi {row['Name']},\n\n{feedback}{footer}\n\nRegards,\nInternship Coordinator",
        )
        print(f"Feedback sent to {row['Name']}")
    save_df(df)


if __name__ == "__main__":
    command = sys.argv[1] if len(sys.argv) > 1 else ""
    if command == "init":
        cmd_init()
    elif command == "remind":
        cmd_remind()
    elif command == "review":
        cmd_review()
    elif command == "all":
        cmd_remind()
        cmd_review()
    else:
        print(__doc__)
