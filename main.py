import sys
from time import sleep

from config import (
	reports_folder,
	report_titles,
	new_report_check_interval_seconds,
	monitor_cycle_wait_seconds,
	latest_update_alert_threshold_minutes,
	latest_update_alert_title,
	DEFAULT_WAIT_MINUTES
)
from utils import (
	extract_latest_update_from_report_body,
	get_latest_outlook_email_report,
	show_os_alert,
	wait_until_minute,
)


def main_monitor(wait_minutes):
	"""Main function to monitor CyberArk reports in Outlook."""

	latest_email_dates = dict.fromkeys(report_titles, None)

	while True:
		reports = {}
		all_reports_are_new = True

		for report_key, report_title in report_titles.items():
			report = get_latest_outlook_email_report(
				title_start=report_title,
				folder_name=reports_folder,
			)

			if report is None:
				all_reports_are_new = False
				break

			received_time = report["received_time"]
			latest_email_date = latest_email_dates[report_key]

			if received_time is None or (
				latest_email_date is not None and received_time <= latest_email_date
			):
				all_reports_are_new = False
				break

			reports[report_key] = report

		if not all_reports_are_new:
			print("Waiting for new emails for all reports...")
			sleep(new_report_check_interval_seconds)
			continue

		print("==== Latest CyberArk Report Updates ===================================")
		alert_messages = []

		for report_key, report in reports.items():
			latest_update = extract_latest_update_from_report_body(report["html_body"])
			latest_email_dates[report_key] = report["received_time"]

			if latest_update is None:
				print(f"{report_key}: no update extracted from {report['received_time']}")
			else:
				print(f"{report_key}: {latest_update} minutes | email date: {report['received_time']}")

				if latest_update > latest_update_alert_threshold_minutes:
					alert_messages.append(
						f"{report_key}: {latest_update} minutes (email date: {report['received_time']})"
					)

		if alert_messages:
			message = (
				"One or more CyberArk reports are above "
				f"{latest_update_alert_threshold_minutes} minutes:\n\n"
				+ "\n".join(alert_messages)
			)
			show_os_alert(latest_update_alert_title, message)

		wait_until_minute(wait_minutes)


if __name__ == "__main__":

	# take param from command line

	wait_minutes = DEFAULT_WAIT_MINUTES

	if len(sys.argv) > 1:
		wait_minutes = sys.argv[1]  # hello

		# check param is integer
		try:
			user_param_int = int(wait_minutes)
			wait_minutes = user_param_int
			print(f"Waiting until minute {wait_minutes} to start monitoring...")
		except ValueError:
			print("Parameter must be an integer.")
			sys.exit(1)

	wait_until_minute(wait_minutes)

	main_monitor(wait_minutes)