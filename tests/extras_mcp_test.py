"""
Outlook Desktop MCP - Tasks/Attachments/OOO/Rules/Categories MCP Test
======================================================================
Tests all new tools through the MCP protocol using the SDK client.
"""
import sys
import os
import json
import asyncio
import logging

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
logging.basicConfig(level=logging.WARNING, stream=sys.stderr)


def log(msg):
    print(msg, file=sys.stderr, flush=True)


async def run_tests():
    from mcp.client.stdio import stdio_client, StdioServerParameters
    from mcp.client.session import ClientSession

    python_exe = r"C:\Development_Local\outlook-desktop-mcp\.venv\Scripts\python.exe"
    server_params = StdioServerParameters(
        command=python_exe,
        args=["-m", "outlook_desktop_mcp.server"],
        cwd=r"C:\Development_Local\outlook-desktop-mcp",
    )

    log("=" * 60)
    log("Outlook Desktop MCP - Extras MCP Test")
    log("=" * 60)
    log("\nConnecting to server...")

    passed = 0
    total = 0
    task_entry_id = None

    async with stdio_client(server_params) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as session:
            await session.initialize()
            log("Server initialized.\n")

            # ----- Test 1: Tool Discovery -----
            total += 1
            log("--- Test 1: Tool Discovery (new tools) ---")
            try:
                tools_result = await session.list_tools()
                tool_names = [t.name for t in tools_result.tools]
                expected = [
                    "list_tasks", "get_task", "create_task", "complete_task",
                    "delete_task", "list_attachments", "save_attachment",
                    "list_categories", "set_category", "list_rules",
                    "toggle_rule", "get_out_of_office",
                ]
                missing = [n for n in expected if n not in tool_names]
                assert not missing, f"Missing tools: {missing}"
                log(f"  All 12 new tools present (total: {len(tool_names)})")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ===== TASKS =====

            total += 1
            log("\n--- Test 2: create_task ---")
            try:
                result = await session.call_tool("create_task", {
                    "subject": "MCP Extras Test Task",
                    "body": "Created through MCP.",
                    "due_date": "2026-03-01",
                    "importance": "high",
                })
                data = json.loads(result.content[0].text)
                task_entry_id = data.get("entry_id")
                log(f"  Created: {data['subject']} (ID: {task_entry_id[:30]}...)")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            total += 1
            log("\n--- Test 3: list_tasks ---")
            try:
                result = await session.call_tool("list_tasks", {"count": 5})
                tasks = json.loads(result.content[0].text)
                log(f"  Got {len(tasks)} tasks:")
                for t in tasks:
                    log(f"    - {t['subject']} ({t['status']})")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            total += 1
            log("\n--- Test 4: get_task ---")
            try:
                assert task_entry_id
                result = await session.call_tool("get_task", {"entry_id": task_entry_id})
                data = json.loads(result.content[0].text)
                log(f"  Subject: {data['subject']}, Importance: {data['importance']}")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            total += 1
            log("\n--- Test 5: complete_task ---")
            try:
                assert task_entry_id
                result = await session.call_tool("complete_task", {"entry_id": task_entry_id})
                log(f"  {result.content[0].text}")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            total += 1
            log("\n--- Test 6: delete_task ---")
            try:
                assert task_entry_id
                result = await session.call_tool("delete_task", {"entry_id": task_entry_id})
                log(f"  {result.content[0].text}")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ===== ATTACHMENTS =====

            total += 1
            log("\n--- Test 7: list_attachments (from inbox email) ---")
            att_entry_id = None
            try:
                # First get an email with attachments
                emails = await session.call_tool("list_emails", {"count": 50})
                email_list = json.loads(emails.content[0].text)
                for e in email_list:
                    if e.get("has_attachments"):
                        att_entry_id = e["entry_id"]
                        break
                if att_entry_id:
                    result = await session.call_tool("list_attachments", {"entry_id": att_entry_id})
                    atts = json.loads(result.content[0].text)
                    log(f"  Found {len(atts)} attachment(s):")
                    for a in atts:
                        log(f"    [{a['index']}] {a['filename']} ({a['size']} bytes)")
                else:
                    log("  No emails with attachments in top 50")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            total += 1
            log("\n--- Test 8: save_attachment ---")
            try:
                if att_entry_id:
                    temp_dir = os.path.join(os.environ.get("TEMP", "/tmp"), "mcp_att_test")
                    result = await session.call_tool("save_attachment", {
                        "entry_id": att_entry_id,
                        "attachment_index": 1,
                        "save_directory": temp_dir,
                    })
                    data = json.loads(result.content[0].text)
                    log(f"  Saved: {data['path']} ({data['size']} bytes)")
                    # Clean up
                    if os.path.exists(data['path']):
                        os.remove(data['path'])
                        log(f"  Cleaned up")
                else:
                    log("  SKIP: No email with attachments found")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ===== CATEGORIES =====

            total += 1
            log("\n--- Test 9: list_categories ---")
            try:
                result = await session.call_tool("list_categories", {})
                cats = json.loads(result.content[0].text)
                log(f"  {len(cats)} categories available (showing first 5):")
                for c in cats[:5]:
                    log(f"    - {c['name']}")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            total += 1
            log("\n--- Test 10: set_category ---")
            try:
                # Get first inbox email
                emails = await session.call_tool("list_emails", {"count": 1})
                email_list = json.loads(emails.content[0].text)
                if email_list:
                    eid = email_list[0]["entry_id"]
                    # Set category
                    result = await session.call_tool("set_category", {
                        "entry_id": eid, "categories": "MCP Test"
                    })
                    log(f"  {result.content[0].text}")
                    # Clear it
                    result = await session.call_tool("set_category", {
                        "entry_id": eid, "categories": ""
                    })
                    log(f"  Cleared: {result.content[0].text}")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ===== RULES =====

            total += 1
            log("\n--- Test 11: list_rules ---")
            try:
                result = await session.call_tool("list_rules", {})
                rules = json.loads(result.content[0].text)
                log(f"  {len(rules)} rules:")
                for r in rules:
                    log(f"    [{r['index']}] {r['name']} (enabled: {r['enabled']})")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ===== OUT OF OFFICE =====

            total += 1
            log("\n--- Test 12: get_out_of_office ---")
            try:
                result = await session.call_tool("get_out_of_office", {})
                data = json.loads(result.content[0].text)
                log(f"  OOF status: {data['status']}")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

    log("")
    log("=" * 60)
    log(f"Results: {passed}/{total} passed")
    log("=" * 60)
    return passed == total


if __name__ == "__main__":
    success = asyncio.run(run_tests())
    sys.exit(0 if success else 1)
