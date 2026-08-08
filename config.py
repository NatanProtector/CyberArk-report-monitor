
# Configuration file for the CyberArk report monitor application.

# Folder in Outlook where the reports are located.
reports_folder = "System Reports"


# Titles of the reports to monitor. These should match the beginning of the email subject lines.
report_title_1 = "CyberArkIncomingStatus "
report_title_2 = "CyberArkOutgoingStatus_Manufacturers"
report_title_3 = "CyberArkOutgoingStatus_Distributers"

# Keys for the different report types.
distributors_key = "distributors"
manufacturers_key = "manufacturers"
incoming_key = "incoming"

# Report titles mapped by report type.
report_titles = {
    incoming_key: report_title_1,
    manufacturers_key: report_title_2,
    distributors_key: report_title_3,
}

# Monitor wait intervals, in seconds.
new_report_check_interval_seconds = 60
monitor_cycle_wait_seconds = 60 * 60

# Regular expression pattern to extract the latest update from the email body.
latest_update_pattern = r'</span><span class="a29">(.*?)</span><span class="a30">'