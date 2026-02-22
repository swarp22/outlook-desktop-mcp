"""
Outlook Desktop MCP Server
===========================
Exposes Microsoft Outlook Desktop (Classic) as an MCP server over stdio.
Uses COM automation — no Microsoft Graph, no Entra app registration.
Just run this on Windows with Outlook open and you have a full email MCP server.

Entry point: python -m outlook_desktop_mcp.server
"""
import sys
import json
import logging

from mcp.server.fastmcp import FastMCP

from outlook_desktop_mcp.com_bridge import OutlookBridge
from datetime import datetime, timedelta

from outlook_desktop_mcp.tools._folder_constants import (
    FOLDER_NAME_TO_ENUM,
    OL_MAIL_ITEM,
    OL_APPOINTMENT_ITEM,
    OL_FOLDER_CALENDAR,
    OL_MEETING,
    OL_MEETING_CANCELED,
    OL_RESPONSE_TENTATIVE,
    OL_RESPONSE_ACCEPTED,
    OL_RESPONSE_DECLINED,
    OL_REQUIRED,
    OL_OPTIONAL,
)
from outlook_desktop_mcp.utils.formatting import (
    format_email_summary,
    format_email_full,
    format_event_summary,
    format_event_full,
)
from outlook_desktop_mcp.utils.errors import format_com_error

# --- Logging (all to stderr, stdout is reserved for MCP JSON-RPC) ---

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("outlook_desktop_mcp")

# --- MCP Server ---

mcp = FastMCP(
    "outlook-desktop-mcp",
    instructions=(
        "This MCP server gives you full access to Microsoft Outlook Desktop on "
        "Windows via COM automation. It can send emails, read inbox messages, "
        "search across folders, mark messages as read/unread, move messages "
        "between folders (including archive), reply to emails, and list the "
        "complete folder hierarchy.\n\n"
        "All operations use the locally authenticated Outlook profile — no "
        "Microsoft Graph API, no Entra app registration, no OAuth tokens needed. "
        "The user's existing Outlook session handles all authentication.\n\n"
        "PREREQUISITE: Outlook Desktop (Classic) must be running. The new/modern "
        "Outlook (olk.exe) is NOT supported — only the classic OUTLOOK.EXE.\n\n"
        "AVAILABLE TOOL CATEGORIES:\n"
        "- Email: send, list, read, search, reply, mark read/unread, move\n"
        "- Calendar: list events, create appointments/meetings, update, delete, "
        "respond to invites, search events\n"
        "- Folders: list folder hierarchy with item counts\n\n"
        "PLANNED (not yet implemented):\n"
        "- Contacts: address book lookups\n"
        "- Tasks: to-do items"
    ),
)

bridge = OutlookBridge()


# --- Helper: resolve folder by name ---

def _resolve_folder(namespace, folder_name: str):
    """Resolve a folder name to an Outlook MAPIFolder object."""
    folder_lower = folder_name.lower().strip()

    if folder_lower in FOLDER_NAME_TO_ENUM:
        return namespace.GetDefaultFolder(FOLDER_NAME_TO_ENUM[folder_lower])

    # Search root folders by name (handles Archive, custom folders)
    root = namespace.DefaultStore.GetRootFolder()
    for i in range(root.Folders.Count):
        f = root.Folders.Item(i + 1)
        if f.Name.lower() == folder_lower:
            return f

    return None


# =====================================================================
# TOOL 1: send_email
# =====================================================================

@mcp.tool()
async def send_email(
    to: str,
    subject: str,
    body: str,
    cc: str = "",
    bcc: str = "",
    html_body: str = "",
) -> str:
    """Send an email using the user's Outlook account.

    Creates and sends an email immediately through the default Outlook profile.
    The email will appear in the user's Sent Items folder after sending.

    Args:
        to: One or more recipient email addresses, separated by semicolons.
            Example: "alice@example.com" or "alice@example.com; bob@example.com"
        subject: The email subject line.
        body: The plain-text body of the email. If html_body is also provided,
            both are set and Outlook will prefer the HTML version.
        cc: Optional. CC recipients, separated by semicolons.
        bcc: Optional. BCC recipients, separated by semicolons.
        html_body: Optional. HTML-formatted body. When provided, Outlook renders
            the email as HTML. The plain-text body serves as fallback.

    Returns:
        A confirmation message with subject and recipients, or an error.
    """
    def _send(outlook, namespace, to, subject, body, cc, bcc, html_body):
        mail = outlook.CreateItem(OL_MAIL_ITEM)
        mail.To = to
        mail.Subject = subject
        mail.Body = body
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc
        if html_body:
            mail.HTMLBody = html_body
        mail.Send()
        return f"Email sent: '{subject}' to {to}"

    try:
        return await bridge.call(_send, to, subject, body, cc, bcc, html_body)
    except Exception as e:
        return f"Error sending email: {format_com_error(e)}"


# =====================================================================
# TOOL 2: list_emails
# =====================================================================

@mcp.tool()
async def list_emails(
    folder: str = "inbox",
    count: int = 10,
    unread_only: bool = False,
) -> str:
    """List recent emails from a specified Outlook folder.

    Returns a JSON array of email summaries sorted by received time (newest
    first). Each summary includes entry_id, subject, sender, sender_name,
    received_time, unread status, and attachment info.

    Use the entry_id from results to read full content with read_email,
    or to perform actions like mark_as_read, move_email, or reply_email.

    Args:
        folder: The folder to list. Case-insensitive names: "inbox" (default),
            "sent"/"sentmail", "drafts", "deleted"/"trash", "junk"/"spam",
            "outbox", "archive", or any custom folder name visible in
            list_folders output.
        count: Maximum number of emails to return. Default 10, max recommended 50.
        unread_only: If true, only return unread emails. Default false.

    Returns:
        JSON array of email summary objects.
    """
    def _list(outlook, namespace, folder, count, unread_only):
        target = _resolve_folder(namespace, folder)
        if not target:
            return json.dumps({"error": f"Folder '{folder}' not found"})

        items = target.Items
        items.Sort("[ReceivedTime]", True)

        if unread_only:
            items = items.Restrict("[UnRead] = True")

        results = []
        limit = min(count, items.Count)
        for i in range(limit):
            try:
                results.append(format_email_summary(items.Item(i + 1)))
            except Exception:
                continue
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, folder, count, unread_only)
    except Exception as e:
        return f"Error listing emails: {format_com_error(e)}"


# =====================================================================
# TOOL 3: read_email
# =====================================================================

@mcp.tool()
async def read_email(
    entry_id: str = "",
    subject_search: str = "",
    folder: str = "inbox",
) -> str:
    """Read the full content of a specific email.

    Retrieves complete email details including body text, recipients, CC,
    and metadata. Provide EITHER entry_id (preferred, exact match) OR
    subject_search (finds most recent match by subject substring).

    Args:
        entry_id: The unique Outlook EntryID of the email. Most reliable way
            to identify a specific email. Get this from list_emails or
            search_emails results.
        subject_search: Alternative to entry_id. A case-insensitive substring
            to search for in email subjects. Returns the most recent match.
        folder: Folder to search when using subject_search. Ignored when
            entry_id is provided. Default "inbox".

    Returns:
        JSON object with full email details (entry_id, subject, sender,
        sender_name, received_time, unread, to, cc, body, attachment info).
    """
    def _read(outlook, namespace, entry_id, subject_search, folder):
        if entry_id:
            item = namespace.GetItemFromID(entry_id)
            return json.dumps(format_email_full(item), indent=2, default=str)

        if not subject_search:
            return json.dumps({"error": "Provide either entry_id or subject_search"})

        target = _resolve_folder(namespace, folder)
        if not target:
            return json.dumps({"error": f"Folder '{folder}' not found"})

        safe_query = subject_search.replace("'", "''")
        filter_str = (
            f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%'"
        )
        items = target.Items.Restrict(filter_str)
        items.Sort("[ReceivedTime]", True)
        if items.Count == 0:
            return json.dumps({"error": f"No email found matching '{subject_search}'"})

        return json.dumps(format_email_full(items.Item(1)), indent=2, default=str)

    try:
        return await bridge.call(_read, entry_id, subject_search, folder)
    except Exception as e:
        return f"Error reading email: {format_com_error(e)}"


# =====================================================================
# TOOL 4: mark_as_read
# =====================================================================

@mcp.tool()
async def mark_as_read(entry_id: str) -> str:
    """Mark a specific email as read in Outlook.

    Changes the unread status to read, same as clicking on an email in Outlook.
    The change is persisted immediately and synced to the server.

    Args:
        entry_id: The unique Outlook EntryID of the email. Get this from
            list_emails or search_emails results.

    Returns:
        Confirmation message with the email subject, or an error.
    """
    def _mark(outlook, namespace, entry_id):
        item = namespace.GetItemFromID(entry_id)
        subject = item.Subject
        item.UnRead = False
        item.Save()
        return f"Marked as read: '{subject}'"

    try:
        return await bridge.call(_mark, entry_id)
    except Exception as e:
        return f"Error marking email as read: {format_com_error(e)}"


# =====================================================================
# TOOL 5: mark_as_unread
# =====================================================================

@mcp.tool()
async def mark_as_unread(entry_id: str) -> str:
    """Mark a specific email as unread in Outlook.

    Restores a previously read email to unread status. Useful for flagging
    emails that need follow-up attention. Persisted immediately.

    Args:
        entry_id: The unique Outlook EntryID of the email. Get this from
            list_emails or search_emails results.

    Returns:
        Confirmation message with the email subject, or an error.
    """
    def _mark(outlook, namespace, entry_id):
        item = namespace.GetItemFromID(entry_id)
        subject = item.Subject
        item.UnRead = True
        item.Save()
        return f"Marked as unread: '{subject}'"

    try:
        return await bridge.call(_mark, entry_id)
    except Exception as e:
        return f"Error marking email as unread: {format_com_error(e)}"


# =====================================================================
# TOOL 6: move_email
# =====================================================================

@mcp.tool()
async def move_email(
    entry_id: str,
    target_folder: str = "archive",
) -> str:
    """Move an email to a different Outlook folder.

    Moves the specified email from its current location to the target folder.
    IMPORTANT: After moving, the email gets a NEW entry_id — the old one
    becomes invalid. Common use: archiving emails after processing.

    Args:
        entry_id: The unique Outlook EntryID of the email to move.
        target_folder: Destination folder name. Default is "archive". Supports
            same names as list_emails: "archive", "inbox", "sent", "deleted"/
            "trash", "drafts", "junk"/"spam", or any custom folder name.

    Returns:
        Confirmation with email subject and destination, or an error.
    """
    def _move(outlook, namespace, entry_id, target_folder):
        item = namespace.GetItemFromID(entry_id)
        subject = item.Subject

        dest = _resolve_folder(namespace, target_folder)
        if not dest:
            return f"Error: Target folder '{target_folder}' not found. Use list_folders to see available folders."

        item.Move(dest)
        return f"Moved '{subject}' to {target_folder}"

    try:
        return await bridge.call(_move, entry_id, target_folder)
    except Exception as e:
        return f"Error moving email: {format_com_error(e)}"


# =====================================================================
# TOOL 7: reply_email
# =====================================================================

@mcp.tool()
async def reply_email(
    entry_id: str,
    body: str,
    reply_all: bool = False,
) -> str:
    """Reply to an email in Outlook.

    Creates and sends a reply, preserving the original message thread.
    Use reply_all=True to reply to all recipients (sender + CC list).

    Args:
        entry_id: The unique Outlook EntryID of the email to reply to.
        body: The reply message text. Prepended above the original message
            in the email thread.
        reply_all: If true, reply to all recipients (sender + all CC/To).
            If false (default), reply only to the sender.

    Returns:
        Confirmation indicating the reply was sent, or an error.
    """
    def _reply(outlook, namespace, entry_id, body, reply_all):
        item = namespace.GetItemFromID(entry_id)
        subject = item.Subject
        reply_item = item.ReplyAll() if reply_all else item.Reply()
        reply_item.Body = body + "\n\n" + reply_item.Body
        reply_item.Send()
        return f"Reply sent to '{subject}' (reply_all={reply_all})"

    try:
        return await bridge.call(_reply, entry_id, body, reply_all)
    except Exception as e:
        return f"Error replying to email: {format_com_error(e)}"


# =====================================================================
# TOOL 8: list_folders
# =====================================================================

@mcp.tool()
async def list_folders(max_depth: int = 2) -> str:
    """List all mail folders in the user's Outlook mailbox.

    Returns a JSON array showing the folder hierarchy with item counts.
    Use this to discover folder names for other tools (list_emails,
    move_email, search_emails). Especially useful for finding the Archive
    folder or any custom user-created folders.

    Args:
        max_depth: How many levels deep to recurse into subfolders.
            Default 2. Set to 1 for top-level only. Max recommended 4.

    Returns:
        JSON array of folder objects with name, item_count, unread_count,
        and subfolders (if any).
    """
    def _list(outlook, namespace, max_depth):
        root = namespace.DefaultStore.GetRootFolder()

        def walk(folder, depth):
            result = {
                "name": folder.Name,
                "item_count": folder.Items.Count,
                "unread_count": folder.UnReadItemCount,
            }
            if depth < max_depth:
                children = []
                for i in range(folder.Folders.Count):
                    try:
                        child = folder.Folders.Item(i + 1)
                        children.append(walk(child, depth + 1))
                    except Exception:
                        continue
                if children:
                    result["subfolders"] = children
            return result

        folders = []
        for i in range(root.Folders.Count):
            f = root.Folders.Item(i + 1)
            folders.append(walk(f, 1))
        return json.dumps(folders, indent=2, default=str)

    try:
        return await bridge.call(_list, max_depth)
    except Exception as e:
        return f"Error listing folders: {format_com_error(e)}"


# =====================================================================
# TOOL 9: search_emails
# =====================================================================

@mcp.tool()
async def search_emails(
    query: str,
    folder: str = "inbox",
    count: int = 10,
) -> str:
    """Search for emails in Outlook using text search.

    Searches email subjects and bodies using Outlook's DASL filter.
    Results are sorted by received time (newest first). Each result
    includes entry_id for further operations.

    Args:
        query: The search term (case-insensitive substring match).
            Examples: "budget report", "meeting notes", "quarterly".
        folder: Folder to search in. Default "inbox". Supports same
            names as list_emails.
        count: Maximum results to return. Default 10.

    Returns:
        JSON array of matching email summaries, or an error.
    """
    def _search(outlook, namespace, query, folder, count):
        target = _resolve_folder(namespace, folder)
        if not target:
            return json.dumps({"error": f"Folder '{folder}' not found"})

        safe_query = query.replace("'", "''")
        filter_str = (
            f"@SQL=("
            f"\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%' OR "
            f"\"urn:schemas:httpmail:textdescription\" LIKE '%{safe_query}%'"
            f")"
        )
        items = target.Items.Restrict(filter_str)
        items.Sort("[ReceivedTime]", True)

        results = []
        limit = min(count, items.Count)
        for i in range(limit):
            try:
                results.append(format_email_summary(items.Item(i + 1)))
            except Exception:
                continue
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_search, query, folder, count)
    except Exception as e:
        return f"Error searching emails: {format_com_error(e)}"


# =====================================================================
# CALENDAR TOOLS
# =====================================================================


# --- Helper: parse ISO date string ---

def _parse_date(date_str: str) -> datetime:
    """Parse ISO 8601 date string like '2026-02-25 14:00' or '2026-02-25T14:00:00'."""
    return datetime.fromisoformat(date_str)


# =====================================================================
# TOOL 10: list_events
# =====================================================================

@mcp.tool()
async def list_events(
    start_date: str = "",
    end_date: str = "",
    count: int = 20,
) -> str:
    """List upcoming calendar events from Outlook.

    Returns a JSON array of event summaries within a date range, sorted by
    start time. Includes recurring event occurrences. Each summary has
    entry_id, subject, start, end, duration, location, organizer, attendees,
    and status info.

    Use entry_id from results with get_event, update_event, delete_event,
    or respond_to_meeting.

    Args:
        start_date: Start of date range in ISO 8601 format (e.g. "2026-02-25"
            or "2026-02-25 09:00"). Default: now.
        end_date: End of date range. Default: 7 days from start_date.
        count: Maximum number of events to return. Default 20.

    Returns:
        JSON array of event summary objects.
    """
    def _list(outlook, namespace, start_date, end_date, count):
        calendar = namespace.GetDefaultFolder(OL_FOLDER_CALENDAR)
        items = calendar.Items

        # CRITICAL ORDER: Sort BEFORE IncludeRecurrences BEFORE Restrict
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        start = _parse_date(start_date) if start_date else datetime.now()
        end = _parse_date(end_date) if end_date else start + timedelta(days=7)

        restrict = (
            f"[Start] >= '{start.strftime('%m/%d/%Y %H:%M')}' "
            f"AND [Start] <= '{end.strftime('%m/%d/%Y %H:%M')}'"
        )
        filtered = items.Restrict(restrict)

        results = []
        n = 0
        for item in filtered:
            n += 1
            try:
                results.append(format_event_summary(item))
            except Exception:
                continue
            if n >= count:
                break

        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, start_date, end_date, count)
    except Exception as e:
        return f"Error listing events: {format_com_error(e)}"


# =====================================================================
# TOOL 11: get_event
# =====================================================================

@mcp.tool()
async def get_event(entry_id: str) -> str:
    """Read the full details of a specific calendar event.

    Retrieves complete event information including body/description,
    attendees, recurrence status, reminders, and response status.

    Args:
        entry_id: The unique Outlook EntryID of the event. Get this from
            list_events or search_events results.

    Returns:
        JSON object with full event details.
    """
    def _get(outlook, namespace, entry_id):
        item = namespace.GetItemFromID(entry_id)
        return json.dumps(format_event_full(item), indent=2, default=str)

    try:
        return await bridge.call(_get, entry_id)
    except Exception as e:
        return f"Error reading event: {format_com_error(e)}"


# =====================================================================
# TOOL 12: create_event
# =====================================================================

@mcp.tool()
async def create_event(
    subject: str,
    start: str,
    end: str,
    location: str = "",
    body: str = "",
    all_day: bool = False,
    reminder_minutes: int = 15,
) -> str:
    """Create a personal calendar appointment (no attendees).

    Creates and saves an appointment on the user's calendar. This is a
    personal event — no meeting invitations are sent. Use create_meeting
    instead if you need to invite attendees.

    Args:
        subject: The event title.
        start: Start time in ISO 8601 format. Examples: "2026-02-25 14:00",
            "2026-02-25T14:00:00". For all-day events, use just the date:
            "2026-02-25".
        end: End time in ISO 8601 format. For all-day events, use the next
            day: "2026-02-26".
        location: Optional. Event location (e.g. "Conference Room A",
            "Microsoft Teams Meeting").
        body: Optional. Description or notes for the event.
        all_day: If true, creates an all-day event. Default false.
        reminder_minutes: Minutes before the event to show a reminder.
            Default 15. Set to 0 to disable reminder.

    Returns:
        Confirmation with event subject and entry_id, or an error.
    """
    def _create(outlook, namespace, subject, start, end, location, body,
                all_day, reminder_minutes):
        appt = outlook.CreateItem(OL_APPOINTMENT_ITEM)
        appt.Subject = subject
        appt.Start = start
        appt.End = end
        if location:
            appt.Location = location
        if body:
            appt.Body = body
        appt.AllDayEvent = all_day
        if reminder_minutes > 0:
            appt.ReminderSet = True
            appt.ReminderMinutesBeforeStart = reminder_minutes
        else:
            appt.ReminderSet = False
        appt.Save()
        return json.dumps({
            "status": "created",
            "subject": appt.Subject,
            "start": str(appt.Start),
            "end": str(appt.End),
            "entry_id": appt.EntryID,
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _create, subject, start, end, location, body, all_day,
            reminder_minutes,
        )
    except Exception as e:
        return f"Error creating event: {format_com_error(e)}"


# =====================================================================
# TOOL 13: create_meeting
# =====================================================================

@mcp.tool()
async def create_meeting(
    subject: str,
    start: str,
    end: str,
    required_attendees: str,
    location: str = "",
    body: str = "",
    optional_attendees: str = "",
) -> str:
    """Create a meeting and send invitations to attendees.

    Creates a calendar meeting and immediately sends meeting requests to
    all specified attendees. The meeting will appear on the organizer's
    calendar and attendees will receive an invitation they can accept,
    decline, or tentatively accept.

    Args:
        subject: The meeting title.
        start: Start time in ISO 8601 format (e.g. "2026-02-25 14:00").
        end: End time in ISO 8601 format (e.g. "2026-02-25 15:00").
        required_attendees: Required attendee email addresses, separated by
            semicolons. Example: "alice@example.com; bob@example.com"
        location: Optional. Meeting location (e.g. "Teams", "Room 301").
        body: Optional. Meeting description or agenda.
        optional_attendees: Optional. Optional attendee emails, separated
            by semicolons.

    Returns:
        Confirmation that the meeting was created and invitations sent.
    """
    def _create(outlook, namespace, subject, start, end, required_attendees,
                location, body, optional_attendees):
        appt = outlook.CreateItem(OL_APPOINTMENT_ITEM)
        appt.Subject = subject
        appt.Start = start
        appt.End = end
        appt.MeetingStatus = OL_MEETING
        if location:
            appt.Location = location
        if body:
            appt.Body = body

        for addr in required_attendees.split(";"):
            addr = addr.strip()
            if addr:
                recip = appt.Recipients.Add(addr)
                recip.Type = OL_REQUIRED

        if optional_attendees:
            for addr in optional_attendees.split(";"):
                addr = addr.strip()
                if addr:
                    recip = appt.Recipients.Add(addr)
                    recip.Type = OL_OPTIONAL

        appt.Recipients.ResolveAll()
        appt.Send()
        return (
            f"Meeting '{subject}' created and invitations sent to "
            f"{required_attendees}"
        )

    try:
        return await bridge.call(
            _create, subject, start, end, required_attendees, location, body,
            optional_attendees,
        )
    except Exception as e:
        return f"Error creating meeting: {format_com_error(e)}"


# =====================================================================
# TOOL 14: update_event
# =====================================================================

@mcp.tool()
async def update_event(
    entry_id: str,
    subject: str = "",
    start: str = "",
    end: str = "",
    location: str = "",
    body: str = "",
) -> str:
    """Update an existing calendar event.

    Modifies properties of an appointment or meeting. Only the fields you
    provide will be updated — omitted fields remain unchanged. For meetings
    you organize, attendees will receive an update notification.

    Args:
        entry_id: The unique Outlook EntryID of the event to update.
        subject: Optional. New event title.
        start: Optional. New start time in ISO 8601 format.
        end: Optional. New end time in ISO 8601 format.
        location: Optional. New location.
        body: Optional. New description/notes.

    Returns:
        Confirmation with updated event details, or an error.
    """
    def _update(outlook, namespace, entry_id, subject, start, end, location, body):
        item = namespace.GetItemFromID(entry_id)
        if subject:
            item.Subject = subject
        if start:
            item.Start = start
        if end:
            item.End = end
        if location:
            item.Location = location
        if body:
            item.Body = body
        item.Save()
        return json.dumps({
            "status": "updated",
            "subject": item.Subject,
            "start": str(item.Start),
            "end": str(item.End),
            "location": item.Location or "",
            "entry_id": item.EntryID,
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _update, entry_id, subject, start, end, location, body,
        )
    except Exception as e:
        return f"Error updating event: {format_com_error(e)}"


# =====================================================================
# TOOL 15: delete_event
# =====================================================================

@mcp.tool()
async def delete_event(entry_id: str) -> str:
    """Delete a calendar event or cancel a meeting.

    For personal appointments, the event is simply deleted. For meetings
    you organized, this cancels the meeting and sends cancellation notices
    to all attendees. For meetings you received, this declines and removes
    the event from your calendar.

    Args:
        entry_id: The unique Outlook EntryID of the event to delete/cancel.

    Returns:
        Confirmation with the event subject, or an error.
    """
    def _delete(outlook, namespace, entry_id):
        item = namespace.GetItemFromID(entry_id)
        subject = item.Subject
        meeting_status = item.MeetingStatus

        # If this is a meeting we organized, cancel it (sends notices)
        if meeting_status == OL_MEETING:
            item.MeetingStatus = OL_MEETING_CANCELED
            item.Send()
            return f"Meeting canceled: '{subject}' (cancellation sent to attendees)"

        # Otherwise just delete
        item.Delete()
        return f"Event deleted: '{subject}'"

    try:
        return await bridge.call(_delete, entry_id)
    except Exception as e:
        return f"Error deleting event: {format_com_error(e)}"


# =====================================================================
# TOOL 16: respond_to_meeting
# =====================================================================

@mcp.tool()
async def respond_to_meeting(
    entry_id: str,
    response: str,
) -> str:
    """Respond to a meeting invitation (accept, decline, or tentative).

    Sends your response to the meeting organizer. The meeting will be
    added to (or updated on) your calendar accordingly.

    Args:
        entry_id: The unique Outlook EntryID of the meeting to respond to.
            Get this from list_events or search_events.
        response: Your response. Must be one of: "accept", "decline",
            or "tentative".

    Returns:
        Confirmation of your response, or an error.
    """
    def _respond(outlook, namespace, entry_id, response):
        response_map = {
            "accept": OL_RESPONSE_ACCEPTED,
            "decline": OL_RESPONSE_DECLINED,
            "tentative": OL_RESPONSE_TENTATIVE,
        }
        response_lower = response.lower().strip()
        if response_lower not in response_map:
            return f"Error: response must be 'accept', 'decline', or 'tentative'. Got: '{response}'"

        item = namespace.GetItemFromID(entry_id)
        subject = item.Subject
        response_item = item.Respond(response_map[response_lower])
        response_item.Send()
        return f"Responded '{response_lower}' to meeting: '{subject}'"

    try:
        return await bridge.call(_respond, entry_id, response)
    except Exception as e:
        return f"Error responding to meeting: {format_com_error(e)}"


# =====================================================================
# TOOL 17: search_events
# =====================================================================

@mcp.tool()
async def search_events(
    query: str,
    start_date: str = "",
    end_date: str = "",
    count: int = 10,
) -> str:
    """Search for calendar events by keyword.

    Searches event subjects within a date range. Results are sorted by
    start time. Includes recurring event occurrences.

    Args:
        query: The search term (case-insensitive substring match on subject).
            Examples: "standup", "review", "1:1".
        start_date: Start of search range in ISO 8601 format. Default: 30
            days ago.
        end_date: End of search range. Default: 30 days from now.
        count: Maximum results to return. Default 10.

    Returns:
        JSON array of matching event summaries.
    """
    def _search(outlook, namespace, query, start_date, end_date, count):
        calendar = namespace.GetDefaultFolder(OL_FOLDER_CALENDAR)
        items = calendar.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        start = _parse_date(start_date) if start_date else datetime.now() - timedelta(days=30)
        end = _parse_date(end_date) if end_date else datetime.now() + timedelta(days=30)

        restrict = (
            f"[Start] >= '{start.strftime('%m/%d/%Y %H:%M')}' "
            f"AND [Start] <= '{end.strftime('%m/%d/%Y %H:%M')}'"
        )
        filtered = items.Restrict(restrict)

        query_lower = query.lower()
        results = []
        for item in filtered:
            if query_lower in (item.Subject or "").lower():
                try:
                    results.append(format_event_summary(item))
                except Exception:
                    continue
                if len(results) >= count:
                    break

        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_search, query, start_date, end_date, count)
    except Exception as e:
        return f"Error searching events: {format_com_error(e)}"


# =====================================================================
# Entry point
# =====================================================================

def main():
    logger.info("Starting Outlook Desktop MCP server...")
    bridge.start()
    logger.info("COM bridge ready. Starting MCP stdio transport...")
    try:
        mcp.run(transport="stdio")
    finally:
        bridge.stop()


if __name__ == "__main__":
    main()
