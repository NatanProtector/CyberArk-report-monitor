import re
from config import latest_update_pattern


def parse_int_from_string(value: str) -> int:
    """
    Safely parse an integer from a string.
    Raises ValueError if the string is not a valid integer.
    """
    if not isinstance(value, str):
        raise TypeError("Input must be a string.")

    value = value.strip()

    try:
        return int(value)
    except ValueError:
        raise ValueError(f"Cannot convert '{value}' to an integer.")


def get_latest_update_from_body_html(body_html: str | None) -> str | None:
    if not body_html:
        return None

    # Get text between </span><span class="a29"> and </span><span class="a30">
    match = re.search(
        latest_update_pattern,
        body_html,
        re.DOTALL,
    )
    if match:
        return match.group(1).strip()
    return None


def get_latest_outlook_email_report(
    title_start: str,
    folder_name: str,
) -> dict | None:
    if not title_start:
        raise ValueError("title_start must not be empty")

    if not folder_name:
        raise ValueError("folder_name must not be empty")

    try:
        import win32com.client
    except ImportError as error:
        raise RuntimeError(
            "pywin32 is required. Install it with: pip install pywin32"
        ) from error

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    folder = _find_outlook_folder(namespace.Folders, folder_name)

    if folder is None:
        raise ValueError(f"Outlook folder was not found: {folder_name}")

    latest_email = _find_latest_email_by_subject_start(folder, title_start)

    if latest_email is None:
        return None

    return {
        "html_body": getattr(latest_email, "HTMLBody", ""),
        "received_time": getattr(latest_email, "ReceivedTime", None),
    }


def get_email_body_html_from_outlook_email(
    title_start: str,
    folder_name: str,
) -> str | None:
    report = get_latest_outlook_email_report(
        title_start=title_start,
        folder_name=folder_name,
    )

    if report is None:
        return None

    return report["html_body"]


def extract_latest_update_from_report_body(
    html_body: str | None,
) -> int | None:

    latest_update = get_latest_update_from_body_html(html_body)

    if latest_update is None:
        return None

    latest_update_int = parse_int_from_string(latest_update)

    return latest_update_int


def show_os_alert(title: str, message: str) -> None:
    try:
        import winsound

        winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
    except RuntimeError:
        pass

    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(0, message, title, 0x00000030)
    except AttributeError as error:
        raise RuntimeError("OS alert windows are only supported on Windows.") from error


def _find_outlook_folder(folders, folder_name: str):
    for folder in folders:
        if folder.Name.casefold() == folder_name.casefold():
            return folder

        nested_folder = _find_outlook_folder(folder.Folders, folder_name)
        if nested_folder is not None:
            return nested_folder

    return None


def _find_latest_email_by_subject_start(folder, title_start: str):
    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    title_start_normalized = title_start.casefold()

    for item in items:
        if getattr(item, "Class", None) != 43:
            continue

        subject = getattr(item, "Subject", "") or ""
        if subject.casefold().startswith(title_start_normalized):
            return item

    return None

