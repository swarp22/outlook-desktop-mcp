"""
Outlook Desktop MCP - Calendar MCP Test
=========================================
Uses the MCP SDK client to test all calendar tools through the MCP protocol.
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
    log("Outlook Desktop MCP - Calendar MCP Test")
    log("=" * 60)
    log("\nConnecting to server...")

    passed = 0
    total = 0
    created_entry_id = None

    async with stdio_client(server_params) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as session:
            await session.initialize()
            log("Server initialized.\n")

            # ----- Test 1: Tool Discovery -----
            total += 1
            log("--- Test 1: Tool Discovery (calendar tools) ---")
            try:
                tools_result = await session.list_tools()
                tool_names = [t.name for t in tools_result.tools]
                calendar_tools = [
                    "list_events", "get_event", "create_event",
                    "create_meeting", "update_event", "delete_event",
                    "respond_to_meeting", "search_events",
                ]
                missing = [n for n in calendar_tools if n not in tool_names]
                assert not missing, f"Missing calendar tools: {missing}"
                log(f"  All 8 calendar tools present (total tools: {len(tool_names)})")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 2: list_events -----
            total += 1
            log("\n--- Test 2: list_events (next 7 days) ---")
            try:
                result = await session.call_tool("list_events", {"count": 5})
                content = result.content[0].text
                events = json.loads(content)
                log(f"  Got {len(events)} events:")
                for e in events:
                    log(f"    - {e['subject'][:50]} ({e['start'][:16]})")
                assert len(events) > 0, "No events returned"
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 3: create_event -----
            total += 1
            log("\n--- Test 3: create_event ---")
            try:
                from datetime import datetime, timedelta
                start = (datetime.now() + timedelta(days=3, hours=2))
                end = start + timedelta(hours=1)
                result = await session.call_tool("create_event", {
                    "subject": "MCP Calendar Test Event",
                    "start": start.strftime("%Y-%m-%d %H:%M"),
                    "end": end.strftime("%Y-%m-%d %H:%M"),
                    "location": "Test Room via MCP",
                    "body": "Created through the MCP calendar tools.",
                    "reminder_minutes": 0,
                })
                content = result.content[0].text
                data = json.loads(content)
                created_entry_id = data.get("entry_id")
                log(f"  Created: {data['subject']}")
                log(f"  EntryID: {created_entry_id[:40]}...")
                assert created_entry_id, "No entry_id returned"
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 4: get_event -----
            total += 1
            log("\n--- Test 4: get_event ---")
            try:
                assert created_entry_id, "No entry_id from Test 3"
                result = await session.call_tool("get_event", {
                    "entry_id": created_entry_id,
                })
                content = result.content[0].text
                data = json.loads(content)
                log(f"  Subject: {data['subject']}")
                log(f"  Location: {data['location']}")
                log(f"  Body: {data['body'][:60]}...")
                assert data["subject"] == "MCP Calendar Test Event"
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 5: create_meeting -----
            total += 1
            log("\n--- Test 5: create_meeting ---")
            try:
                start = (datetime.now() + timedelta(days=4, hours=3))
                end = start + timedelta(minutes=30)
                result = await session.call_tool("create_meeting", {
                    "subject": "MCP Calendar Meeting Test",
                    "start": start.strftime("%Y-%m-%d %H:%M"),
                    "end": end.strftime("%Y-%m-%d %H:%M"),
                    "required_attendees": "aaanerud@microsoft.com",
                    "location": "Virtual via MCP",
                    "body": "Meeting created through MCP calendar tools.",
                })
                content = result.content[0].text
                log(f"  Result: {content}")
                assert "sent" in content.lower() or "created" in content.lower()
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 6: update_event -----
            total += 1
            log("\n--- Test 6: update_event ---")
            try:
                assert created_entry_id, "No entry_id from Test 3"
                result = await session.call_tool("update_event", {
                    "entry_id": created_entry_id,
                    "subject": "MCP Calendar Test Event (UPDATED)",
                    "location": "Updated Room via MCP",
                })
                content = result.content[0].text
                data = json.loads(content)
                log(f"  Updated: {data['subject']}")
                log(f"  Location: {data['location']}")
                assert "UPDATED" in data["subject"]
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 7: search_events -----
            total += 1
            log("\n--- Test 7: search_events ---")
            try:
                result = await session.call_tool("search_events", {
                    "query": "MCP Calendar",
                    "count": 5,
                })
                content = result.content[0].text
                results = json.loads(content)
                log(f"  Found {len(results)} events matching 'MCP Calendar'")
                for r in results:
                    log(f"    - {r['subject'][:50]}")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 8: delete_event -----
            total += 1
            log("\n--- Test 8: delete_event ---")
            try:
                assert created_entry_id, "No entry_id from Test 3"
                result = await session.call_tool("delete_event", {
                    "entry_id": created_entry_id,
                })
                content = result.content[0].text
                log(f"  Result: {content}")
                assert "deleted" in content.lower()
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
