
# Configuration file for the CyberArk report monitor application.

# Folder in Outlook where the reports are located.
reports_folder = "System Reports"


# Titles of the reports to monitor. These should match the beginning of the email subject lines.
report_title_1 = "CyberArkIncomingStatus "
report_title_2 = "CyberArkOutgoingStatus_Manufacturers"
report_title_3 = "CyberArkOutgoingStatus_Distributers"

# Regular expression pattern to extract the latest update from the email body.
latest_update_pattern = r'</span><span class="a29">(.*?)</span><span class="a30">'