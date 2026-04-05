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
import re

from mcp.server.fastmcp import FastMCP

from outlook_desktop_mcp.com_bridge import OutlookBridge
from datetime import datetime, timedelta

import os

from outlook_desktop_mcp.tools._folder_constants import (
    FOLDER_NAME_TO_ENUM,
    OL_FOLDER_CALENDAR,
    OL_FOLDER_TASKS,
)
from outlook_desktop_mcp.utils.formatting import (
    format_email_summary,
    format_email_full,
    format_event_summary,
    format_event_full,
    format_task_summary,
    format_task_full,
)
from outlook_desktop_mcp.utils.errors import format_com_error

# --- Logging (all to stderr, stdout is reserved for MCP JSON-RPC) ---

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("outlook_desktop_mcp")


# --- Security helpers ---

def _safe_dasl(query: str) -> str:
    """Sanitize a string for use in a DASL LIKE filter value.
    Escapes SQL wildcards (% and _) so user input is treated as literals,
    then escapes quote characters required by DASL syntax.
    """
    query = query.replace("%", "[%]").replace("_", "[_]")
    return query.replace("'", "''").replace('"', '""')


# --- MCP Server ---

mcp = FastMCP(
    "outlook-desktop-mcp",
    instructions=(
        "This MCP server gives you READ-ONLY access to Microsoft Outlook Desktop on "
        "Windows via COM automation. It can read inbox messages, "
        "search across folders, and list the complete folder hierarchy.\n\n"
        "All operations use the locally authenticated Outlook profile — no "
        "Microsoft Graph API, no Entra app registration, no OAuth tokens needed. "
        "The user's existing Outlook session handles all authentication.\n\n"
        "PREREQUISITE: Outlook Desktop (Classic) must be running. The new/modern "
        "Outlook (olk.exe) is NOT supported — only the classic OUTLOOK.EXE.\n\n"
        "IMPORTANT: This server is READ-ONLY. It cannot send, create, modify, "
        "or delete any items in Outlook.\n\n"
        "AVAILABLE TOOL CATEGORIES:\n"
        "- Email: list, read, search, attachments\n"
        "- Calendar: list events, get event details, search events\n"
        "- Tasks: list, get task details\n"
        "- Categories: list color categories\n"
        "- Rules: list mail rules\n"
        "- Out of Office: check auto-reply status\n"
        "- Folders: list folder hierarchy with item counts"
    ),
)

bridge = OutlookBridge()


# --- Helper: resolve store by account name ---

def _resolve_store(namespace, account: str = ""):
    """Resolve an account name to an Outlook Store object.

    If account is empty, returns DefaultStore.
    Otherwise does a case-insensitive substring match on Store.DisplayName.
    """
    if not account:
        return namespace.DefaultStore

    account_lower = account.lower().strip()
    for i in range(namespace.Stores.Count):
        store = namespace.Stores.Item(i + 1)
        if account_lower in store.DisplayName.lower():
            return store

    return None


def _require_store(namespace, account: str = ""):
    """Resolve store, raising ValueError if not found."""
    store = _resolve_store(namespace, account)
    if store is None:
        raise ValueError(f"Account '{account}' not found. Use list_accounts to see available accounts.")
    return store


# --- Helper: resolve folder by name ---

def _walk_folders(parent, name_lower: str):
    """Recursively search subfolders of parent for a folder matching name_lower."""
    for i in range(parent.Folders.Count):
        try:
            f = parent.Folders.Item(i + 1)
            if f.Name.lower() == name_lower:
                return f
            found = _walk_folders(f, name_lower)
            if found:
                return found
        except Exception:
            continue
    return None


def _resolve_folder(namespace, folder_name: str, store=None):
    """Resolve a folder name to an Outlook MAPIFolder object.

    Resolution order:
    1. Slash-delimited path (e.g. "Inbox/Receipts") — traverse segment by segment
    2. Built-in Outlook folder enum (inbox, sent, deleted, etc.)
    3. Root-level folder name match (fast path)
    4. Recursive depth-first search of entire folder tree (fallback)
    """
    folder_name = folder_name.strip()
    store = store or namespace.DefaultStore

    # Slash-delimited path: traverse segment by segment
    if "/" in folder_name:
        parts = [p.strip() for p in folder_name.split("/")]
        current = _resolve_folder(namespace, parts[0], store)
        if current is None:
            return None
        for part in parts[1:]:
            part_lower = part.lower()
            found = None
            for i in range(current.Folders.Count):
                try:
                    f = current.Folders.Item(i + 1)
                    if f.Name.lower() == part_lower:
                        found = f
                        break
                except Exception:
                    continue
            if found is None:
                return None
            current = found
        return current

    folder_lower = folder_name.lower()

    # Built-in Outlook folders
    if folder_lower in FOLDER_NAME_TO_ENUM:
        return store.GetDefaultFolder(FOLDER_NAME_TO_ENUM[folder_lower])

    # Root-level search (fast path)
    root = store.GetRootFolder()
    for i in range(root.Folders.Count):
        try:
            f = root.Folders.Item(i + 1)
            if f.Name.lower() == folder_lower:
                return f
        except Exception:
            continue

    # Recursive fallback: search entire folder tree
    return _walk_folders(root, folder_lower)


# =====================================================================
# TOOL: list_accounts
# =====================================================================

@mcp.tool()
async def list_accounts() -> str:
    """List all Outlook accounts (stores) configured in the profile.

    Returns a JSON array of account objects with display_name, store_id,
    and is_default. Use the display_name (or a unique substring) as the
    'account' parameter in other tools to target a specific account.

    Returns:
        JSON array of account objects.
    """
    def _list(outlook, namespace):
        default_id = namespace.DefaultStore.StoreID
        results = []
        for i in range(namespace.Stores.Count):
            store = namespace.Stores.Item(i + 1)
            results.append({
                "display_name": store.DisplayName,
                "store_id": store.StoreID,
                "is_default": store.StoreID == default_id,
            })
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list)
    except Exception as e:
        return f"Error listing accounts: {format_com_error(e)}"


# =====================================================================
# TOOL: list_emails
# =====================================================================

@mcp.tool()
async def list_emails(
    folder: str = "inbox",
    count: int = 10,
    unread_only: bool = False,
    start_date: str = "",
    end_date: str = "",
    account: str = "",
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
        start_date: Optional. Only return emails received on or after this date.
            ISO 8601 format (e.g. "2026-03-10" or "2026-03-10 09:00").
        end_date: Optional. Only return emails received on or before this date.
            ISO 8601 format. Default: now (if start_date is provided).
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of email summary objects.
    """
    def _list(outlook, namespace, folder, count, unread_only, start_date, end_date, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            return json.dumps({"error": f"Folder '{folder}' not found"})

        items = target.Items
        items.Sort("[ReceivedTime]", True)

        # Build restriction filters
        restrictions = []
        if unread_only:
            restrictions.append("[UnRead] = True")
        if start_date:
            start = _parse_date(start_date)
            restrictions.append(f"[ReceivedTime] >= '{start.strftime('%m/%d/%Y %H:%M')}'")
        if end_date:
            end = _parse_date(end_date)
            restrictions.append(f"[ReceivedTime] <= '{end.strftime('%m/%d/%Y %H:%M')}'")
        elif start_date:
            # Default end to now when start is specified
            restrictions.append(f"[ReceivedTime] <= '{datetime.now().strftime('%m/%d/%Y %H:%M')}'")

        if restrictions:
            items = items.Restrict(" AND ".join(restrictions))

        results = []
        limit = min(count, items.Count)
        for i in range(limit):
            try:
                results.append(format_email_summary(items.Item(i + 1)))
            except Exception:
                continue
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, folder, count, unread_only, start_date, end_date, account)
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
    account: str = "",
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
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON object with full email details (entry_id, subject, sender,
        sender_name, received_time, unread, to, cc, body, attachment info).
    """
    def _read(outlook, namespace, entry_id, subject_search, folder, account):
        if entry_id:
            item = namespace.GetItemFromID(entry_id)
            return json.dumps(format_email_full(item), indent=2, default=str)

        if not subject_search:
            return json.dumps({"error": "Provide either entry_id or subject_search"})

        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            return json.dumps({"error": f"Folder '{folder}' not found"})

        safe_query = _safe_dasl(subject_search)
        filter_str = (
            f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%'"
        )
        items = target.Items.Restrict(filter_str)
        items.Sort("[ReceivedTime]", True)
        if items.Count == 0:
            return json.dumps({"error": f"No email found matching '{subject_search}'"})

        return json.dumps(format_email_full(items.Item(1)), indent=2, default=str)

    try:
        return await bridge.call(_read, entry_id, subject_search, folder, account)
    except Exception as e:
        return f"Error reading email: {format_com_error(e)}"


# =====================================================================
# TOOL: list_folders
# =====================================================================

@mcp.tool()
async def list_folders(folder: str = "", max_depth: int = 3, account: str = "") -> str:
    """List mail folders in the user's Outlook mailbox.

    When called with no folder argument, lists top-level folders. Provide a
    folder name to drill into its subfolders — use this to browse the full
    folder tree step by step (e.g. first call with no folder to see top-level,
    then call with folder="Inbox" to see Inbox children, then
    folder="Inbox/Projects" to go deeper).

    Folder names from this output can be used directly in list_emails,
    move_email, search_emails, etc. Use slash-delimited paths for nested
    folders (e.g. "Inbox/Receipts/2026").

    Args:
        folder: Optional. Folder to list children of. Supports folder names
            ("Inbox"), slash paths ("Inbox/Receipts"), or built-in names
            ("sent", "drafts"). When empty, lists from the mailbox root.
        max_depth: How many levels deep to recurse below the starting folder.
            Default 3. Set to 1 to see only immediate children.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of folder objects with name, full_path, item_count,
        unread_count, and subfolders (if any).
    """
    def _list(outlook, namespace, folder, max_depth, account):
        max_depth = min(max(1, max_depth), 10)
        store = _require_store(namespace, account)

        if folder:
            start = _resolve_folder(namespace, folder, store)
            if not start:
                return json.dumps({"error": f"Folder '{folder}' not found"})
            base_path = folder
        else:
            start = store.GetRootFolder()
            base_path = ""

        def walk(f, depth, path_prefix):
            current_path = f"{path_prefix}/{f.Name}" if path_prefix else f.Name
            result = {
                "name": f.Name,
                "full_path": current_path,
                "item_count": f.Items.Count,
                "unread_count": f.UnReadItemCount,
            }
            if depth < max_depth:
                children = []
                for i in range(f.Folders.Count):
                    try:
                        child = f.Folders.Item(i + 1)
                        children.append(walk(child, depth + 1, current_path))
                    except Exception:
                        continue
                if children:
                    result["subfolders"] = children
            return result

        folders = []
        for i in range(start.Folders.Count):
            try:
                child = start.Folders.Item(i + 1)
                folders.append(walk(child, 1, base_path))
            except Exception:
                continue
        return json.dumps(folders, indent=2, default=str)

    try:
        return await bridge.call(_list, folder, max_depth, account)
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
    start_date: str = "",
    end_date: str = "",
    account: str = "",
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
        start_date: Optional. Only return emails received on or after this date.
            ISO 8601 format (e.g. "2026-03-10" or "2026-03-10 09:00").
        end_date: Optional. Only return emails received on or before this date.
            ISO 8601 format. Default: now (if start_date is provided).
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of matching email summaries, or an error.
    """
    def _search(outlook, namespace, query, folder, count, start_date, end_date, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            return json.dumps({"error": f"Folder '{folder}' not found"})

        safe_query = _safe_dasl(query)
        dasl_parts = [
            f"(\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%' OR "
            f"\"urn:schemas:httpmail:textdescription\" LIKE '%{safe_query}%')"
        ]
        if start_date:
            start = _parse_date(start_date)
            dasl_parts.append(
                f"\"urn:schemas:httpmail:datereceived\" >= '{start.strftime('%m/%d/%Y %H:%M')}'"
            )
        if end_date:
            end = _parse_date(end_date)
            dasl_parts.append(
                f"\"urn:schemas:httpmail:datereceived\" <= '{end.strftime('%m/%d/%Y %H:%M')}'"
            )
        elif start_date:
            dasl_parts.append(
                f"\"urn:schemas:httpmail:datereceived\" <= '{datetime.now().strftime('%m/%d/%Y %H:%M')}'"
            )

        filter_str = "@SQL=" + " AND ".join(dasl_parts)
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
        return await bridge.call(_search, query, folder, count, start_date, end_date, account)
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
    account: str = "",
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
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of event summary objects.
    """
    def _list(outlook, namespace, start_date, end_date, count, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        calendar = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
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
        return await bridge.call(_list, start_date, end_date, count, account)
    except Exception as e:
        return f"Error listing events: {format_com_error(e)}"


# =====================================================================
# TOOL 11: get_event
# =====================================================================

@mcp.tool()
async def get_event(entry_id: str, account: str = "") -> str:
    """Read the full details of a specific calendar event.

    Retrieves complete event information including body/description,
    attendees, recurrence status, reminders, and response status.

    Args:
        entry_id: The unique Outlook EntryID of the event. Get this from
            list_events or search_events results.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON object with full event details.
    """
    def _get(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        return json.dumps(format_event_full(item), indent=2, default=str)

    try:
        return await bridge.call(_get, entry_id, account)
    except Exception as e:
        return f"Error reading event: {format_com_error(e)}"


# =====================================================================
# TOOL: search_events
# =====================================================================

@mcp.tool()
async def search_events(
    query: str,
    start_date: str = "",
    end_date: str = "",
    count: int = 10,
    account: str = "",
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
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of matching event summaries.
    """
    def _search(outlook, namespace, query, start_date, end_date, count, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        calendar = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
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
        return await bridge.call(_search, query, start_date, end_date, count, account)
    except Exception as e:
        return f"Error searching events: {format_com_error(e)}"


# =====================================================================
# TASK TOOLS
# =====================================================================

@mcp.tool()
async def list_tasks(
    include_completed: bool = False,
    count: int = 20,
    account: str = "",
) -> str:
    """List tasks from the Outlook Tasks folder.

    Returns a JSON array of task summaries sorted by due date. Each task
    includes entry_id, subject, status, percent_complete, due_date,
    importance, and categories.

    Args:
        include_completed: If true, include completed tasks. Default false
            (only pending/in-progress tasks).
        count: Maximum number of tasks to return. Default 20.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of task summary objects.
    """
    def _list(outlook, namespace, include_completed, count, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        folder = store.GetDefaultFolder(OL_FOLDER_TASKS)
        items = folder.Items
        items.Sort("[DueDate]")

        if not include_completed:
            items = items.Restrict("[Complete] = False")

        results = []
        limit = min(count, items.Count)
        for i in range(limit):
            try:
                results.append(format_task_summary(items.Item(i + 1)))
            except Exception:
                continue
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, include_completed, count, account)
    except Exception as e:
        return f"Error listing tasks: {format_com_error(e)}"


@mcp.tool()
async def get_task(entry_id: str, account: str = "") -> str:
    """Read the full details of a specific task.

    Args:
        entry_id: The unique Outlook EntryID of the task.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON object with full task details including body.
    """
    def _get(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        return json.dumps(format_task_full(item), indent=2, default=str)

    try:
        return await bridge.call(_get, entry_id, account)
    except Exception as e:
        return f"Error reading task: {format_com_error(e)}"


# =====================================================================
# ATTACHMENT TOOLS
# =====================================================================

@mcp.tool()
async def list_attachments(entry_id: str, account: str = "") -> str:
    """List all attachments on an email or calendar event.

    Args:
        entry_id: The EntryID of the email or event to check for attachments.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON array of attachment objects with index, filename, and size.
    """
    def _list(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        results = []
        for i in range(item.Attachments.Count):
            att = item.Attachments.Item(i + 1)
            results.append({
                "index": i + 1,
                "filename": att.FileName,
                "size": att.Size,
            })
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, entry_id, account)
    except Exception as e:
        return f"Error listing attachments: {format_com_error(e)}"


@mcp.tool()
async def save_attachment(
    entry_id: str,
    attachment_index: int = 1,
    save_directory: str = "",
    account: str = "",
) -> str:
    """Save an attachment from an email or event to disk.

    Downloads the specified attachment to a local directory.

    Args:
        entry_id: The EntryID of the email or event containing the attachment.
        attachment_index: Which attachment to save (1-based index). Default 1
            (first attachment). Use list_attachments to see available indices.
        save_directory: Directory to save the file to. Default: user's
            Downloads folder.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        The full file path where the attachment was saved, or an error.
    """
    def _save(outlook, namespace, entry_id, attachment_index, save_directory, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if attachment_index < 1 or item.Attachments.Count < attachment_index:
            return f"Error: Only {item.Attachments.Count} attachment(s), requested index {attachment_index}"

        att = item.Attachments.Item(attachment_index)
        if not save_directory:
            save_directory = os.path.join(os.path.expanduser("~"), "Downloads")

        # Resolve to real path before creating
        save_directory = os.path.realpath(save_directory)
        os.makedirs(save_directory, exist_ok=True)

        # Strip path separators and dangerous characters from filename
        safe_name = os.path.basename(att.FileName)
        safe_name = re.sub(r'[^\w\.\-_ ]', '_', safe_name)
        if not safe_name:
            safe_name = "attachment"

        save_path = os.path.join(save_directory, safe_name)

        # Ensure final path is still inside the intended directory
        if not os.path.realpath(save_path).startswith(save_directory + os.sep) and \
           os.path.realpath(save_path) != save_directory:
            return "Error: Attachment filename would escape the target directory."

        att.SaveAsFile(save_path)
        return json.dumps({
            "status": "saved",
            "filename": safe_name,
            "path": save_path,
            "size": att.Size,
        }, indent=2, default=str)

    try:
        return await bridge.call(_save, entry_id, attachment_index, save_directory, account)
    except Exception as e:
        return f"Error saving attachment: {format_com_error(e)}"


# =====================================================================
# CATEGORY TOOLS
# =====================================================================

@mcp.tool()
async def list_categories(account: str = "") -> str:
    """List all available Outlook categories.

    Returns the color categories configured in the user's Outlook profile.
    These can be applied to emails, events, tasks, and other items.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of category objects with name and color index.
    """
    def _list(outlook, namespace, account):
        # Categories are profile-wide, not per-store, but we accept the param for consistency
        results = []
        for i in range(namespace.Categories.Count):
            cat = namespace.Categories.Item(i + 1)
            results.append({"name": cat.Name, "color": cat.Color})
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, account)
    except Exception as e:
        return f"Error listing categories: {format_com_error(e)}"


# =====================================================================
# RULES TOOLS
# =====================================================================

@mcp.tool()
async def list_rules(account: str = "") -> str:
    """List all mail rules in Outlook.

    Returns the configured inbox rules with their names and enabled status.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of rule objects with name, enabled status, and index.
    """
    def _list(outlook, namespace, account):
        store = _require_store(namespace, account)
        rules = store.GetRules()
        results = []
        for i in range(rules.Count):
            rule = rules.Item(i + 1)
            results.append({
                "index": i + 1,
                "name": rule.Name,
                "enabled": bool(rule.Enabled),
            })
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, account)
    except Exception as e:
        return f"Error listing rules: {format_com_error(e)}"


# =====================================================================
# OUT OF OFFICE TOOLS
# =====================================================================

@mcp.tool()
async def get_out_of_office(account: str = "") -> str:
    """Check the current Out of Office (auto-reply) status.

    Returns whether Out of Office is currently enabled.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON object with the OOF status.
    """
    def _get(outlook, namespace, account):
        store = _require_store(namespace, account)
        try:
            prop_tag = "http://schemas.microsoft.com/mapi/proptag/0x661D000B"
            oof_state = store.PropertyAccessor.GetProperty(prop_tag)
            return json.dumps({
                "out_of_office": bool(oof_state),
                "status": "on" if oof_state else "off",
            }, indent=2)
        except Exception:
            return json.dumps({
                "out_of_office": None,
                "status": "unknown",
                "note": "Could not read OOF property. Check Outlook settings directly.",
            }, indent=2)

    try:
        return await bridge.call(_get, account)
    except Exception as e:
        return f"Error checking OOF status: {format_com_error(e)}"


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
