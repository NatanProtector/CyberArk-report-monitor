


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

select_text = "options:\n 1 - start\n 2 - delayed start\n 3 - help\n\nselect option:"

user_input = input(select_text)

while True:

    if user_input == "1":
        wait_minutes = 0
        break
    elif user_input == "2":

        minutes_input = input("Enter minute of the hour (0-59): ")
        try:
            minutes_ = int(minutes_input)
            if minutes_ < 0 or minutes_ > 59:
                raise ValueError("Invalid input.")
        except ValueError:
            print("Invalid input. Using default time.")
            minutes_ = DEFAULT_WAIT_MINUTES

        wait_minutes = minutes_
        break

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
            1 - Start now; later cycles run at :{DEFAULT_WAIT_MINUTES:02d} each hour (default).
            2 - Enter a minute (0-59); start at its next occurrence, then hourly.
                Example: enter 15 at 6:10 to start at 6:15 (5 minutes later).
            3 - Show this guide and exit; reopen the app to start.
            Keep the window open. Ctrl+C stops the app. Close alerts to continue.

            CHECKS
            Starts with the latest email of each type; then waits for all types to be newer.
            Retries every {new_report_check_interval_seconds} seconds while waiting.
            Reports need the expected HTML format with a whole-number update age.

            Need different settings or report formats? Contact the app maintainer.
            Settings are built into the EXE. More details: README.md.
            """)

    else:
        print("Invalid option. Please select 1, 2, or 3.")
        user_input = input(select_text)     


wait_until_minute(wait_minutes, display=True)

main_monitor(wait_minutes)