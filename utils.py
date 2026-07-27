def parse_int_from_string(s):
    """
    Safely parse an integer from a string.
    Raises ValueError if the string is not a valid integer.
    """
    if not isinstance(s, str):
        raise TypeError("Input must be a string.")

    s = s.strip()  # Remove leading/trailing spaces

    try:
        return int(s)  # Convert to integer
    except ValueError:
        raise ValueError(f"Cannot convert '{s}' to an integer.")

try:
    num = parse_int_from_string("  42  ")
    print("Parsed integer:", num)  # Output: 42
except (TypeError, ValueError) as e:
    print("Error:", e)


def get_latest_update_from_body_html(body_html: str) -> str | None:
    # Get text between </span><span class="a29"> and </span><span class="a30">
    match = re.search(
        r'</span><span class="a29">(.*?)</span><span class="a30">',
        body_html,
        re.DOTALL,
    )
    if match:
        return match.group(1).strip()
    return None

def get_email_body_html_from_outlook_email(
	title_start: str,
    folder_name: str,
) -> str:
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

    html_body = getattr(latest_email, "HTMLBody", "")
    
    return html_body

def extract_latest_update_from_report_body(
    html_body: str,
) -> str | None:

    latest_update = get_latest_update_from_body_html(html_body)

    if latest_update is None:
        return None

    latest_update_int = parse_int_from_string(latest_update)

    return latest_update_int


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

