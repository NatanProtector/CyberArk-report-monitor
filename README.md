# CyberArk Report Monitor

A Windows app that reads CyberArk report emails in classic Outlook. It shows
each report's update age and email received time, then sounds an alert and opens
a dialog if an age is **over 35 minutes** (default).

## Prerequisites

- Windows and **classic Outlook**, signed in to the report mailbox under the
  same Windows account. New Outlook and Outlook on the web are not supported.
- Existing report emails in the expected HTML format with a whole-number update
  age. The app does not generate or deliver reports.

## Outlook Setup

Put all three report types directly in **System Reports**:

```text
Your mailbox
+-- Inbox
    +-- System Reports
        +-- Incoming report emails
        +-- Manufacturers report emails
        +-- Distributors report emails
```

The folder can be anywhere in Outlook. Names are matched case-insensitively;
the first matching folder is used, so avoid duplicates. Subfolders are not read.
Use an Outlook rule to move reports into it automatically.

Subjects must **start with** these values (case-insensitive; extra text may follow):

| Report | Subject prefix (shown in quotes) |
| --- | --- |
| Incoming | `"CyberArkIncomingStatus "` |
| Manufacturers | `"CyberArkOutgoingStatus_Manufacturers"` |
| Distributors | `"CyberArkOutgoingStatus_Distributers"` |

Omit the quotes, keep the incoming prefix's **trailing space**, and preserve
the spelling **Distributers**.

## Run the App

Launch `CyberArkReportMonitor.exe`. No additional runtime installation is needed.

| Option | What happens |
| --- | --- |
| **1 - Start** | Starts monitoring now. Later cycles use the default minute (`:52`) each hour. |
| **2 - Delayed start** | Enter a minute (`0-59`). Starts at its next occurrence, then checks at that minute each hour. |
| **3 - Help** | Shows the guide and exits. Reopen the app to start. |

For example, enter **15** at **6:10** to start at **6:15**, a **5-minute wait**.
The input is a clock minute, not a delay. If that minute has already started,
the next occurrence is in the following hour.

Keep the window open; **Ctrl+C** stops the app. **Close alert dialogs** to continue.

## Monitoring Behavior

- Starts with the latest email of each type, including emails received before startup.
- Then waits for **all three types** to have newer emails, retrying every **60 seconds**.
  Missing emails alone do not trigger alerts.
- Uses the age inside the report, not the email's age. Exactly 35 minutes does not alert.
- After processing and closing alerts, waits for the next scheduled minute.
  Restarting clears the record of previously processed emails.

## Configuration

Settings are built into the EXE. Ask the maintainer for a rebuilt version to change them.

| Setting | Purpose | Default |
| --- | --- | --- |
| `reports_folder` / `report_titles` | Folder / subject prefixes | See Outlook Setup |
| `new_report_check_interval_seconds` | Retry interval | `60` seconds |
| `latest_update_alert_threshold_minutes` | Alert above this age | `35` minutes |
| `latest_update_alert_title` | Alert title | `CyberArk Report Alert` |
| `latest_update_pattern` | HTML extraction pattern | See below |
| `DEFAULT_WAIT_MINUTES` | Option 1 cycle minute / fallback for invalid option 2 input | `52` |

Maintainers: edit [config.py](config.py). The pattern captures an integer between
`</span><span class="a29">` and `</span><span class="a30">`.
`monitor_cycle_wait_seconds` is unused; scheduling uses the selected minute.

## Build an EXE With PyInstaller (Maintainers)

Build on Windows from the project directory (not needed by EXE users):

```powershell
py -3 -m venv venv
.\venv\Scripts\python.exe -m pip install pywin32 pyinstaller
.\venv\Scripts\python.exe -m PyInstaller --onefile --console --name CyberArkReportMonitor --hidden-import win32timezone start.py
```

Output: `dist\CyberArkReportMonitor.exe`. Keep `--console` for the menu and results.
Dependencies are bundled; Outlook and the mailbox setup are still required.
Rebuild after configuration changes; a separate config file beside the EXE has
no effect. Test on the target Windows/Outlook installation before distributing.

## Troubleshooting

| Message | What to check |
| --- | --- |
| Folder not found | Folder name and mailbox access in classic Outlook. |
| Waiting for new emails | All three subjects match; newer emails are directly in the folder. |
| No update extracted | Maintainer needs to check the HTML format and extraction pattern. |
| Integer conversion error | Report age must be a whole number; invalid values stop the app. |
| Missing dependency | Ask the maintainer for a rebuilt EXE with all dependencies. |

## Project Files

[start.py](start.py): menu and guide; [main.py](main.py): monitoring and scheduled
startup; [config.py](config.py): settings; [utils.py](utils.py): Outlook, extraction,
alerts, and scheduling.
