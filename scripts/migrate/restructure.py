import ast
import pathlib
import re
import shutil

OLD = pathlib.Path("word_document_server")
NEW = pathlib.Path("word_mcp")

TABLES = {"word_live_add_table", "word_live_format_table", "word_live_modify_table"}
REFERENCES = {
    "word_live_insert_image",
    "word_live_insert_cross_reference",
    "word_live_list_cross_reference_items",
    "word_live_insert_equation",
}
SOURCES = {
    "live_tools": OLD / "tools" / "live_tools.py",
    "live_read_tools": OLD / "tools" / "live_read_tools.py",
    "live_layout_tools": OLD / "tools" / "live_layout_tools.py",
    "screen_capture_tools": OLD / "tools" / "screen_capture_tools.py",
}
SINGLE_TARGET = {
    "live_read_tools": "read",
    "live_layout_tools": "layout",
    "screen_capture_tools": "screen",
}
IMPORT_REWRITES = [
    ("word_document_server.core.word_com", "word_mcp.word_com"),
    ("word_document_server.core.table_com", "word_mcp.table_com"),
    ("from word_document_server.core import table_com", "from word_mcp import table_com"),
    ("word_document_server.utils.text_safety", "word_mcp.text_safety"),
    ("word_document_server.defaults", "word_mcp.defaults"),
]


def strip_mac(text):
    out = []
    skipping = False
    for line in text.splitlines():
        if skipping:
            if line.startswith('    if sys.platform != "win32":'):
                skipping = False
                out.append(line)
            elif re.match(r"(async )?def ", line):
                raise SystemExit(f"mac block ran into a def: {line}")
            continue
        if line == "    if _MAC_AVAILABLE:":
            skipping = True
            continue
        if "_MAC_AVAILABLE" in line or line.strip() == "# macOS JXA dispatch":
            continue
        out.append(line)
    if skipping:
        raise SystemExit("unterminated mac block")
    return "\n".join(out) + "\n"


def rewrite_imports(text):
    for old, new in IMPORT_REWRITES:
        text = text.replace(old, new)
    if "word_document_server" in text:
        raise SystemExit("unhandled word_document_server import")
    return text


def target_for(module, name):
    if module != "live_tools":
        return SINGLE_TARGET[module]
    if name in TABLES:
        return "tables"
    if name in REFERENCES:
        return "references"
    return "edit"


def split_live_tools(text):
    lines = text.splitlines()
    defs = [
        n for n in ast.parse(text).body
        if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))
    ]
    starts = []
    floor = 0
    for node in defs:
        start = node.lineno - 1
        while start > floor and (
            lines[start - 1].strip() == "" or lines[start - 1].lstrip().startswith("#")
        ):
            start -= 1
        starts.append(start)
        floor = node.end_lineno
    header = "\n".join(lines[: starts[0]]).rstrip() + "\n"
    parts = {"edit": [], "tables": [], "references": []}
    for i, node in enumerate(defs):
        end = starts[i + 1] if i + 1 < len(defs) else len(lines)
        block = "\n".join(lines[starts[i]:end]).strip("\n")
        parts[target_for("live_tools", node.name)].append(block)
    return {name: header + "\n\n" + "\n\n\n".join(blocks) + "\n" for name, blocks in parts.items()}


def tool_names(text):
    return [
        n.name for n in ast.parse(text).body
        if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)) and not n.name.startswith("_")
    ]


def build_tools():
    (NEW / "tools").mkdir(parents=True, exist_ok=True)
    (NEW / "__init__.py").write_text("", encoding="utf-8")
    (NEW / "tools" / "__init__.py").write_text("", encoding="utf-8")
    owner = {}
    for module, path in SOURCES.items():
        text = rewrite_imports(strip_mac(path.read_text(encoding="utf-8")))
        outputs = split_live_tools(text) if module == "live_tools" else {SINGLE_TARGET[module]: text}
        for target, body in outputs.items():
            ast.parse(body)
            (NEW / "tools" / f"{target}.py").write_text(body, encoding="utf-8", newline="\n")
            for name in tool_names(body):
                owner[name] = target
    return owner


def build_server(owner):
    source = (OLD / "main.py").read_text(encoding="utf-8")
    lines = source.splitlines()
    register = next(
        n for n in ast.parse(source).body
        if isinstance(n, ast.FunctionDef) and n.name == "register_tools"
    )
    pattern = re.compile(r"\b(" + "|".join(SOURCES) + r")\.(\w+)")
    blocks = []
    for node in register.body:
        if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        start = min([d.lineno for d in node.decorator_list] + [node.lineno])
        text = "\n".join(lines[start - 1: node.end_lineno])
        if not pattern.search(text):
            continue
        text = pattern.sub(lambda m: f"{owner[m.group(2)]}.{m.group(2)}", text)
        text = re.sub(r"^(\s*)def ", r"\1async def ", text, count=1, flags=re.M)
        text, awaited = re.subn(
            r"\breturn (edit|tables|references|read|layout|screen)\.", r"return await \1.", text
        )
        if awaited != 1:
            raise SystemExit(f"{node.name}: expected one delegating return, found {awaited}")
        blocks.append(text)
    if len(blocks) != 45:
        raise SystemExit(f"expected 45 live wrappers, found {len(blocks)}")
    body = "\n\n".join(blocks)
    defaults = [n for n in ("DEFAULT_AUTHOR", "DEFAULT_INITIALS") if re.search(rf"\b{n}\b", body)]
    defaults_import = f"from word_mcp.defaults import {', '.join(defaults)}\n" if defaults else ""
    server = (
        "import os\n"
        "import sys\n"
        "\n"
        'os.environ.setdefault("FASTMCP_LOG_LEVEL", "WARNING")\n'
        "\n"
        "from fastmcp import FastMCP\n"
        "from mcp.types import ToolAnnotations\n"
        "\n"
        f"{defaults_import}"
        "from word_mcp.tools import edit, layout, read, references, screen, tables\n"
        "\n"
        'mcp = FastMCP("word-mcp")\n'
        "\n"
        "\n"
        "def register_tools():\n"
        f"{body}\n"
        "\n"
        "\n"
        "def run():\n"
        '    if sys.platform != "win32":\n'
        '        print("word-mcp only runs on Windows", file=sys.stderr)\n'
        "        raise SystemExit(1)\n"
        "    register_tools()\n"
        '    mcp.run(transport="stdio", show_banner=False)\n'
    )
    ast.parse(server)
    (NEW / "server.py").write_text(server, encoding="utf-8", newline="\n")


def move_support():
    shutil.copy(OLD / "defaults.py", NEW / "defaults.py")
    shutil.copy(OLD / "utils" / "text_safety.py", NEW / "text_safety.py")
    shutil.copy(OLD / "core" / "table_com.py", NEW / "table_com.py")
    shutil.copy(OLD / "core" / "word_com.py", NEW / "word_com.py")
    for name in ("defaults.py", "text_safety.py", "table_com.py", "word_com.py"):
        path = NEW / name
        path.write_text(rewrite_imports(path.read_text(encoding="utf-8")), encoding="utf-8", newline="\n")


owner = build_tools()
build_server(owner)
move_support()
print(f"{len(owner)} public functions, {len(set(owner.values()))} tool modules")
