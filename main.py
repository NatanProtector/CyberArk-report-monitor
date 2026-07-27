from utils import get_email_body_html_from_outlook_email, extract_latest_update_from_report_body 
from config import reports_folder, report_title_1, report_title_2, report_title_3

if __name__ == "__main__":

	print("==== Extracting Latest Outlook Email Update ===================================")

	for report_title in [report_title_1, report_title_2, report_title_3]:
		html_body = get_email_body_html_from_outlook_email(
			title_start=report_title,
			folder_name=reports_folder,
		)
          
		latest_update = extract_latest_update_from_report_body(html_body)

		if latest_update is None:
			print("No matching email found or no update extracted.")

		else:
			print("Report Title:", report_title)
			print("Latest Update:", latest_update, "minutes")