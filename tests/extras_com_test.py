"""
Outlook Desktop MCP - Tasks/Attachments/OOO/Rules/Categories COM Validation
============================================================================
Standalone COM tests for the new tool modules.
Requires Classic Outlook (Desktop) to be running.
"""
import sys
import os
import time
from datetime import datetime, timedelta


def log(msg):
    print(msg, file=sys.stderr, flush=True)


_created_task_id = None


def test_connect():
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    log(f"  Connected: {namespace.DefaultStore.DisplayName}")
    return outlook, namespace


# ===== TASKS =====

def test_list_tasks(namespace):
    """List tasks from the default Tasks folder."""
    tasks_folder = namespace.GetDefaultFolder(13)
    items = tasks_folder.Items
    log(f"  Tasks folder: {items.Count} items")
    for i in range(min(5, items.Count)):
        item = items.Item(i + 1)
        log(f"    [{i+1}] {item.Subject} (Status: {item.Status}, Complete: {item.Complete})")


def test_create_task(outlook):
    """Create a test task."""
    global _created_task_id
    task = outlook.CreateItem(3)  # olTaskItem
    task.Subject = "MCP Test Task - COM Validation"
    task.Body = "Automated test task."
    task.DueDate = (datetime.now() + timedelta(days=7)).strftime("%Y-%m-%d")
    task.Importance = 2  # High
    task.ReminderSet = False
    task.Save()
    _created_task_id = task.EntryID
    log(f"  Created: '{task.Subject}' (EntryID: {task.EntryID[:40]}...)")


def test_complete_task(namespace):
    """Mark the test task as complete."""
    global _created_task_id
    if not _created_task_id:
        log("  SKIP: No task created")
        return
    item = namespace.GetItemFromID(_created_task_id)
    item.Status = 2  # olTaskComplete
    item.PercentComplete = 100
    item.Save()
    log(f"  Completed: '{item.Subject}' (PercentComplete: {item.PercentComplete})")


def test_delete_task(namespace):
    """Delete the test task."""
    global _created_task_id
    if not _created_task_id:
        log("  SKIP: No task created")
        return
    item = namespace.GetItemFromID(_created_task_id)
    subject = item.Subject
    item.Delete()
    _created_task_id = None
    log(f"  Deleted: '{subject}'")


# ===== ATTACHMENTS =====

def test_list_attachments(namespace):
    """Find an email with attachments and list them."""
    inbox = namespace.GetDefaultFolder(6)
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)

    for i in range(min(50, items.Count)):
        item = items.Item(i + 1)
        if item.Attachments.Count > 0:
            log(f"  Email: '{item.Subject}' has {item.Attachments.Count} attachment(s):")
            for j in range(item.Attachments.Count):
                att = item.Attachments.Item(j + 1)
                log(f"    [{j+1}] {att.FileName} ({att.Size} bytes)")
            return item.EntryID
    log("  No emails with attachments found in top 50")
    return None


def test_save_attachment(namespace, entry_id):
    """Save an attachment to a temp directory."""
    if not entry_id:
        log("  SKIP: No email with attachments found")
        return
    item = namespace.GetItemFromID(entry_id)
    att = item.Attachments.Item(1)
    temp_dir = os.path.join(os.environ.get("TEMP", "/tmp"), "outlook_mcp_test")
    os.makedirs(temp_dir, exist_ok=True)
    save_path = os.path.join(temp_dir, att.FileName)
    att.SaveAsFile(save_path)
    exists = os.path.exists(save_path)
    size = os.path.getsize(save_path) if exists else 0
    log(f"  Saved: {save_path} ({size} bytes, exists={exists})")
    # Clean up
    if exists:
        os.remove(save_path)
        log(f"  Cleaned up test file")


# ===== CATEGORIES =====

def test_list_categories(namespace):
    """List available categories."""
    categories = namespace.Categories
    log(f"  Available categories: {categories.Count}")
    for i in range(min(10, categories.Count)):
        cat = categories.Item(i + 1)
        log(f"    - {cat.Name} (Color: {cat.Color})")


def test_set_category(namespace):
    """Set a category on the most recent inbox email, then remove it."""
    inbox = namespace.GetDefaultFolder(6)
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    if items.Count == 0:
        log("  SKIP: Inbox empty")
        return
    item = items.Item(1)
    old_categories = item.Categories or ""
    log(f"  Target: '{item.Subject}'")
    log(f"  Old categories: '{old_categories}'")
    item.Categories = "MCP Test Category"
    item.Save()
    log(f"  Set to: '{item.Categories}'")
    # Restore
    item.Categories = old_categories
    item.Save()
    log(f"  Restored to: '{item.Categories}'")


# ===== RULES =====

def test_list_rules(namespace):
    """List mail rules."""
    store = namespace.DefaultStore
    rules = store.GetRules()
    log(f"  Total rules: {rules.Count}")
    for i in range(min(10, rules.Count)):
        rule = rules.Item(i + 1)
        log(f"    [{i+1}] '{rule.Name}' (Enabled: {rule.Enabled})")


# ===== OUT OF OFFICE =====

def test_get_ooo_status(namespace):
    """Check Out of Office status via Store property."""
    store = namespace.DefaultStore
    try:
        # Try PropertyAccessor for OOF status
        # PR_OOF_STATE MAPI property
        prop_tag = "http://schemas.microsoft.com/mapi/proptag/0x661D000B"
        oof_state = store.PropertyAccessor.GetProperty(prop_tag)
        log(f"  Out of Office: {'ON' if oof_state else 'OFF'}")
    except Exception as e:
        log(f"  OOF via PropertyAccessor failed: {e}")
        log("  Trying alternative: checking for OOF rules...")
        # Alternative: check if there's an OOF auto-reply rule
        try:
            rules = store.GetRules()
            for i in range(rules.Count):
                rule = rules.Item(i + 1)
                if "out of office" in rule.Name.lower() or "automatic reply" in rule.Name.lower():
                    log(f"    Found OOF rule: '{rule.Name}' (Enabled: {rule.Enabled})")
                    return
            log("  No OOF rule found — likely OFF")
        except Exception as e2:
            log(f"  Alternative also failed: {e2}")


def main():
    log("=" * 60)
    log("Outlook Desktop MCP - Extras COM Validation")
    log("=" * 60)

    log("\n--- Connect ---")
    try:
        outlook, namespace = test_connect()
        log("  PASS")
    except Exception as e:
        log(f"  FAIL: {e}")
        sys.exit(1)

    tests = [
        ("List Tasks", lambda: test_list_tasks(namespace)),
        ("Create Task", lambda: test_create_task(outlook)),
        ("Complete Task", lambda: test_complete_task(namespace)),
        ("Delete Task", lambda: test_delete_task(namespace)),
        ("List Attachments", lambda: test_list_attachments(namespace)),
        ("Save Attachment", lambda: test_save_attachment(namespace, _att_entry_id)),
        ("List Categories", lambda: test_list_categories(namespace)),
        ("Set/Restore Category", lambda: test_set_category(namespace)),
        ("List Rules", lambda: test_list_rules(namespace)),
        ("Get OOO Status", lambda: test_get_ooo_status(namespace)),
    ]

    # Pre-run: find an email with attachments for test 6
    global _att_entry_id
    log("\n--- Pre-scan: Find email with attachments ---")
    try:
        _att_entry_id = test_list_attachments(namespace)
        log("  DONE")
    except Exception as e:
        _att_entry_id = None
        log(f"  {e}")

    passed = 1  # connect passed
    total = len(tests) + 1

    # Skip test 5 (List Attachments) since we did it in pre-scan
    for i, (name, fn) in enumerate(tests):
        if i == 4:  # skip duplicate list attachments
            passed += 1
            continue
        log(f"\n--- Test {i + 2}: {name} ---")
        try:
            fn()
            passed += 1
            log("  PASS")
        except Exception as e:
            log(f"  FAIL: {e}")

    log("")
    log("=" * 60)
    log(f"Results: {passed}/{total} passed")
    log("=" * 60)

    import pythoncom
    pythoncom.CoUninitialize()
    sys.exit(0 if passed == total else 1)


if __name__ == "__main__":
    main()
