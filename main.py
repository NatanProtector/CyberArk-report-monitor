from __future__ import annotations
import re
from utils import get_email_body_html_from_outlook_email, extract_latest_update_from_report_body


if __name__ == "__main__":

	folder = "System Reports"

	report_title_1 = "CyberArkIncomingStatus "
	report_title_2 = "CyberArkOutgoingStatus_Manufacturers"
	report_title_3 = "CyberArkOutgoingStatus_Distributers"

	print("==== Extracting Latest Outlook Email Update ===================================")

	for report_title in [report_title_1, report_title_2, report_title_3]:
		html_body = get_email_body_html_from_outlook_email(
			title_start=report_title,
			folder_name=folder,
		)
          
		latest_update = extract_latest_update_from_report_body(html_body)

		if latest_update is None:
			print("No matching email found or no update extracted.")

		else:
			print("Report Title:", report_title)
			print("Latest Update:", latest_update, "minutes")