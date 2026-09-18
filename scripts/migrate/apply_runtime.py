import ast
import pathlib
import re

TOOLS = pathlib.Path("word_mcp/tools")
READ_ONLY = {
    "word_live_list_cross_reference_items",
    "word_live_take_snapshot",
    "word_live_get_diff",
    "word_live_snapshot_status",
    "word_live_get_text",
    "word_live_get_paragraph_format",
    "word_live_get_info",
    "word_live_list_open",
    "word_live_find_text",
    "word_live_get_comments",
    "word_live_list_revisions",
    "word_live_get_page_text",
    "word_live_get_undo_history",
    "word_live_diagnose_layout",
    "word_screen_capture",
}
OLD_HANDLER = '    except Exception as e:\n        return json.dumps({"error": str(e)})'
NEW_HANDLER = "    except Exception as e:\n        return error_json(e)"
HEADER = "from word_mcp.com_runtime import error_json\nfrom word_mcp.live_tool import live_tool\n"

total = 0
for path in sorted(TOOLS.glob("*.py")):
    if path.name == "__init__.py":
        continue
    text = path.read_text(encoding="utf-8")
    names = re.findall(r"^async def (\w+)", text, flags=re.M)
    handlers = text.count(OLD_HANDLER)
    if handlers != len(names):
        raise SystemExit(f"{path}: {len(names)} tools but {handlers} blanket handlers")
    text = text.replace(OLD_HANDLER, NEW_HANDLER)
    text = re.sub(
        r"^async def (\w+)",
        lambda m: f"@live_tool(mutates={m.group(1) not in READ_ONLY})\ndef {m.group(1)}",
        text,
        flags=re.M,
    )
    text = re.sub(r"\bawait (word_\w+)\(", r"\1.sync(", text)
    text = text.replace("from word_mcp.word_com import", "from word_mcp.com_runtime import")
    text, inserted = re.subn(r"^import json\n", "import json\n" + HEADER, text, count=1, flags=re.M)
    if inserted != 1:
        raise SystemExit(f"{path}: no top-level 'import json' to anchor the new imports")
    if "async def" in text or "await " in text:
        raise SystemExit(f"{path}: async code remains")
    ast.parse(text)
    path.write_text(text, encoding="utf-8", newline="\n")
    total += len(names)
    print(f"{path.name}: {len(names)} tools")

if total != 45:
    raise SystemExit(f"expected 45 tools, migrated {total}")
