


# take param from command line

from config import (
    DEFAULT_WAIT_MINUTES,
    latest_update_alert_threshold_minutes,
    new_report_check_interval_seconds,
    reports_folder,
    report_titles,
)
from main import main_monitor
from utils import wait_until_minute

wait_minutes = DEFAULT_WAIT_MINUTES

select_text = "select option: \n1 - start\n 2 - delayed start\n  3 - help"

user_input = input(select_text)

while user_input not in ["1", "2", "3"]:
    print("Invalid option")
    user_input = input(select_text)
if user_input == "1":
    wait_minutes = 0
elif user_input == "2":
    wait_minutes = DEFAULT_WAIT_MINUTES
elif user_input == "3":
    report_subjects = "\n".join(f"  - {title!r}" for title in report_titles.values())
    print(f"""
CyberArk Report Monitor - Quick Guide

Shows report update ages; alerts when an age exceeds
{latest_update_alert_threshold_minutes} minutes.

SETUP
  Windows + classic Outlook, signed in to the report mailbox.
  New Outlook and Outlook on the web are not supported.
  Put all reports directly in {reports_folder!r}, not in subfolders.
  Example: Mailbox > Inbox > {reports_folder}
  Any folder location works; avoid duplicate names. A mail rule can move reports.

  Subjects must start with (case-insensitive; omit quotes, keep spaces):
{report_subjects}

RUN CyberArkReportMonitor.exe
  1 - Start at the next hour (:00), then hourly. Not an immediate start.
  2 - Start at the next :{DEFAULT_WAIT_MINUTES:02d}, then hourly at :{DEFAULT_WAIT_MINUTES:02d}.
  3 - Show this guide and exit; reopen the app to start.
  Keep the window open. Ctrl+C stops the app. Close alerts to continue.

CHECKS
  Starts with the latest email of each type; then waits for all types to be newer.
  Retries every {new_report_check_interval_seconds} seconds while waiting.
  Reports need the expected HTML format with a whole-number update age.

Need different settings or report formats? Contact the app maintainer.
Settings are built into the EXE. More details: README.md.
""")
    exit()


wait_until_minute(wait_minutes, display=True)

main_monitor(wait_minutes)