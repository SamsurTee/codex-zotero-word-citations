#!/usr/bin/env python3
"""High-level Zotero-to-Word workflow for Codex.

This script turns structured literature records plus drafted sections into a
live Word document with real Zotero citation fields and a bibliography.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import re
import subprocess
import sys
import time
import urllib.error
import urllib.request
import uuid
from pathlib import Path
from typing import Any


DEFAULT_STYLE = "APA Style 7th edition"
CITE_PATTERN = re.compile(r"\[\[cite:([^\]]+)\]\]")
DEFAULT_ZOTERO_MCP_PYTHON = Path.home() / ".local/share/uv/tools/zotero-mcp-server/bin/python"


def skill_dir() -> Path:
    return Path(__file__).resolve().parents[1]


def word_plugin_script() -> Path:
    return Path(__file__).resolve().with_name("zotero_word_plugin.py")


def current_python_has_zotero_mcp() -> bool:
    return importlib.util.find_spec("zotero_mcp") is not None


def zotero_mcp_python() -> str:
    override = os.environ.get("ZOTERO_MCP_PYTHON")
    if override:
        candidate = Path(override).expanduser()
        if not candidate.exists():
            raise RuntimeError(f"ZOTERO_MCP_PYTHON does not exist: {candidate}")
        return str(candidate)

    if DEFAULT_ZOTERO_MCP_PYTHON.exists():
        return str(DEFAULT_ZOTERO_MCP_PYTHON)

    if sys.executable and current_python_has_zotero_mcp():
        return sys.executable

    raise RuntimeError(
        "Could not locate a Python runtime with zotero_mcp installed. "
        "Set ZOTERO_MCP_PYTHON or install zotero-mcp-server first."
    )


def import_records_into_zotero(records: list[dict[str, Any]], tags: list[str] | None = None) -> list[dict[str, Any]]:
    tool_python = zotero_mcp_python()
    code = """
import json
import sys
from zotero_mcp.server import _import_records_into_zotero

payload = json.load(sys.stdin)
result = _import_records_into_zotero(
    payload.get("records") or [],
    inherited_tags=payload.get("tags") or None,
)
json.dump(result, sys.stdout, ensure_ascii=False)
"""
    proc = subprocess.run(
        [tool_python, "-c", code],
        input=json.dumps({"records": records, "tags": tags or []}, ensure_ascii=False),
        text=True,
        capture_output=True,
        env={
            **os.environ,
            "ZOTERO_LOCAL": os.environ.get("ZOTERO_LOCAL", "true"),
            "ZOTERO_LOCAL_PORT": os.environ.get("ZOTERO_LOCAL_PORT", "23119"),
        },
    )
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or "zotero-mcp import subprocess failed")

    try:
        payload = json.loads(proc.stdout or "[]")
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Could not parse zotero-mcp import response: {proc.stdout}") from exc

    if not isinstance(payload, list):
        raise RuntimeError(f"Unexpected zotero-mcp import payload: {payload}")
    return payload


def run_command(args: list[str]) -> str:
    proc = subprocess.run(args, text=True, capture_output=True)
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or f"Command failed: {' '.join(args)}")
    return proc.stdout.strip()


def run_osascript(script: str, *argv: str) -> str:
    proc = subprocess.run(
        ["osascript", "-"] + list(argv),
        input=script,
        text=True,
        capture_output=True,
    )
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or "osascript failed")
    return proc.stdout.strip()


def clipboard_contents() -> str:
    return subprocess.run(
        ["pbpaste"],
        text=True,
        capture_output=True,
        check=False,
    ).stdout


def set_clipboard_text(text: str) -> None:
    subprocess.run(["pbcopy"], input=text, text=True, check=True)


def with_temporary_clipboard(text: str):
    class _ClipboardContext:
        def __enter__(self_inner) -> None:
            self_inner.previous = clipboard_contents()
            set_clipboard_text(text)

        def __exit__(self_inner, exc_type, exc, tb) -> None:
            subprocess.run(
                ["pbcopy"],
                input=self_inner.previous,
                text=True,
                check=False,
            )

    return _ClipboardContext()


def ensure_word_document(create_new_document: bool) -> None:
    if create_new_document:
        script = """
tell application "Microsoft Word"
  activate
  make new document
end tell
"""
    else:
        script = """
tell application "Microsoft Word"
  activate
  if not (exists active document) then
    error "No active Word document"
  end if
end tell
"""
    run_osascript(script)
    time.sleep(1.0)


def word_type_text(text: str) -> None:
    if not text:
        return
    script = """
tell application "Microsoft Word"
  activate
end tell
delay 0.1
tell application "System Events"
  keystroke "v" using command down
end tell
"""
    with with_temporary_clipboard(text):
        run_osascript(script)
        time.sleep(0.1)
        word_go_to_document_end()


def word_press_end() -> None:
    script = """
tell application "Microsoft Word"
  activate
end tell
tell application "System Events"
  key code 119
end tell
"""
    run_osascript(script)


def word_newline() -> None:
    script = """
tell application "Microsoft Word"
  activate
end tell
tell application "System Events"
  key code 36
end tell
"""
    run_osascript(script)


def word_go_to_document_end() -> None:
    script = """
tell application "Microsoft Word"
  activate
end tell
tell application "System Events"
  key code 119 using command down
end tell
"""
    run_osascript(script)


def word_clear_document() -> None:
    script = """
tell application "Microsoft Word"
  activate
  if not (exists active document) then
    error "No active Word document"
  end if
end tell
"""
    run_osascript(script)
    select_all = """
tell application "System Events"
  keystroke "a" using command down
  key code 51
end tell
"""
    run_osascript(select_all)


def word_save_document_as(path: Path) -> None:
    path = path.expanduser().resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()

    # Word's AppleScript "save as" handling is unreliable for an already-active
    # document. Use UI automation with clipboard paste for deterministic Save As.
    previous_clipboard = clipboard_contents()
    try:
        run_osascript(
            """
tell application "Microsoft Word"
  activate
end tell
"""
        )
        time.sleep(0.2)
        run_osascript(
            """
tell application "System Events"
  keystroke "s" using {command down, shift down}
end tell
"""
        )
        time.sleep(0.8)
        run_osascript(
            """
tell application "System Events"
  keystroke "g" using {command down, shift down}
end tell
"""
        )
        time.sleep(0.5)
        with with_temporary_clipboard(str(path.parent)):
            run_osascript(
                """
tell application "System Events"
  keystroke "v" using command down
  key code 36
end tell
"""
            )
        time.sleep(0.8)
        with with_temporary_clipboard(path.name):
            run_osascript(
                """
tell application "System Events"
  keystroke "a" using command down
  keystroke "v" using command down
  key code 36
end tell
"""
            )
        deadline = time.time() + 5.0
        while time.time() < deadline:
            if path.exists():
                break
            time.sleep(0.2)
        if not path.exists():
            raise RuntimeError(f"Word Save As did not create the expected file: {path}")
    finally:
        subprocess.run(
            ["pbcopy"],
            input=previous_clipboard,
            text=True,
            check=False,
        )


def call_word_plugin(*args: str) -> dict[str, Any]:
    python_bin = sys.executable or "python3"
    output = run_command([python_bin, str(word_plugin_script()), *args])
    try:
        return json.loads(output) if output else {}
    except json.JSONDecodeError:
        return {"raw": output}


def parse_import_results(imported_items: list[dict[str, Any]]) -> list[dict[str, str]]:
    results: list[dict[str, str]] = []
    for item in imported_items:
        if isinstance(item, dict):
            data = item.get("data", item)
        else:
            data = {}
        record_id = ""
        for tag in data.get("tags", []) or []:
            tag_value = tag.get("tag", "") if isinstance(tag, dict) else str(tag)
            if tag_value.startswith("codex-record:"):
                record_id = tag_value.split(":", 1)[1]
                break
        results.append(
            {
                "record_id": record_id,
                "item_key": item.get("key", data.get("key", "")) if isinstance(item, dict) else data.get("key", ""),
                "title": data.get("title", ""),
                "doi": data.get("DOI", ""),
            }
        )
    return results


def build_item_key_map(imported_results: list[dict[str, str]]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for result in imported_results:
        record_id = result.get("record_id", "").strip()
        item_key = result.get("item_key", "").strip()
        if record_id and item_key:
            mapping[record_id] = item_key
    return mapping


def split_ris_records(ris_text: str) -> list[str]:
    records: list[str] = []
    current: list[str] = []
    for line in ris_text.splitlines():
        current.append(line)
        if line.startswith("ER  -"):
            record = "\n".join(current).strip()
            if record:
                records.append(record + "\n")
            current = []
    if current:
        record = "\n".join(current).strip()
        if record:
            records.append(record + "\n")
    return records


def import_ris_text_into_zotero(ris_text: str, timeout_s: float = 180.0) -> list[dict[str, Any]]:
    port = os.environ.get("ZOTERO_LOCAL_PORT", "23119")
    url = f"http://127.0.0.1:{port}/connector/import?session={uuid.uuid4().hex}"
    req = urllib.request.Request(
        url,
        data=ris_text.encode("utf-8"),
        headers={"Content-Type": "application/x-research-info-systems"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout_s) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(
            f"Zotero RIS import failed (HTTP {exc.code}): {body or exc.reason}"
        ) from exc

    if not isinstance(payload, list):
        raise RuntimeError(f"Unexpected RIS import payload: {payload}")
    return payload


def import_ris_file_in_batches(ris_path: Path, batch_size: int = 5) -> list[dict[str, Any]]:
    if batch_size <= 0:
        raise RuntimeError("batch_size must be greater than 0")
    records = split_ris_records(ris_path.read_text(errors="ignore"))
    if not records:
        raise RuntimeError(f"No RIS records found in {ris_path}")

    imported_items: list[dict[str, Any]] = []
    for start in range(0, len(records), batch_size):
        batch_records = records[start:start + batch_size]
        payload = import_ris_text_into_zotero("\n\n".join(batch_records))
        if len(payload) != len(batch_records):
            raise RuntimeError(
                "RIS batch import returned the wrong number of items "
                f"for records {start + 1}-{start + len(batch_records)}: "
                f"expected {len(batch_records)}, got {len(payload)}"
            )
        imported_items.extend(payload)
    return imported_items


def insert_citation_for_ids(citation_ids: list[str], item_key_map: dict[str, str], style: str) -> None:
    item_keys: list[str] = []
    for citation_id in citation_ids:
        key = item_key_map.get(citation_id)
        if not key:
            raise RuntimeError(f"Missing imported Zotero item for citation id '{citation_id}'")
        item_keys.append(key)

    args = ["insert"]
    for item_key in item_keys:
        args.extend(["--item-key", item_key])
    args.extend(["--style", style])
    call_word_plugin(*args)
    time.sleep(0.8)
    word_go_to_document_end()


def render_text_with_citations(text: str, item_key_map: dict[str, str], style: str) -> None:
    cursor = 0
    for match in CITE_PATTERN.finditer(text):
        plain_text = text[cursor:match.start()]
        word_type_text(plain_text)
        citation_ids = [part.strip() for part in match.group(1).split(",") if part.strip()]
        if not citation_ids:
            raise RuntimeError(f"Empty citation token in text: {match.group(0)}")
        insert_citation_for_ids(citation_ids, item_key_map, style)
        cursor = match.end()
    word_type_text(text[cursor:])


def run_workflow(spec: dict[str, Any]) -> dict[str, Any]:
    records = spec.get("records") or []
    if not isinstance(records, list) or not records:
        raise RuntimeError("Workflow spec must include a non-empty 'records' list")

    sections = spec.get("sections") or []
    if not isinstance(sections, list) or not sections:
        raise RuntimeError("Workflow spec must include a non-empty 'sections' list")

    style = str(spec.get("style") or DEFAULT_STYLE)
    document_title = str(spec.get("document_title") or "Zotero Word Summary")
    create_new_document = bool(spec.get("create_new_document", True))
    global_tags = spec.get("tags") or []
    bibliography_heading = str(spec.get("bibliography_heading") or "References")
    insert_bibliography = bool(spec.get("insert_bibliography", True))
    output_path = str(spec.get("output_path") or "").strip()

    imported_items = import_records_into_zotero(records, tags=global_tags)
    imported_results = parse_import_results(imported_items)
    item_key_map = build_item_key_map(imported_results)

    missing_ids = [
        str(record.get("id") or record.get("record_id"))
        for record in records
        if str(record.get("id") or record.get("record_id")) not in item_key_map
    ]
    if missing_ids:
        raise RuntimeError(
            "Imported items were missing record tags for ids: " + ", ".join(sorted(missing_ids))
        )

    ensure_word_document(create_new_document=create_new_document)
    call_word_plugin("check")

    word_type_text(document_title + "\n\n")
    for section in sections:
        heading = str(section.get("heading") or "").strip()
        text = str(section.get("text") or "")
        if heading:
            word_type_text(heading + "\n")
        render_text_with_citations(text, item_key_map, style)
        word_type_text("\n\n")

    if insert_bibliography:
        if bibliography_heading:
            word_type_text(bibliography_heading + "\n")
        call_word_plugin("bibliography", "--style", style)
        word_type_text("\n")

    call_word_plugin("refresh")
    if output_path:
        word_save_document_as(Path(output_path))

    return {
        "document_title": document_title,
        "style": style,
        "imported_items": imported_results,
        "item_key_map": item_key_map,
        "bibliography_inserted": insert_bibliography,
        "output_path": output_path or None,
    }


def load_json(path: Path) -> Any:
    return json.loads(path.read_text())


def cmd_import_records(args: argparse.Namespace) -> int:
    records = load_json(Path(args.json_file))
    if not isinstance(records, list):
        raise RuntimeError("Input JSON must be an array of records")
    imported_items = import_records_into_zotero(records, tags=args.tag)
    print(json.dumps(parse_import_results(imported_items), ensure_ascii=False, indent=2))
    return 0


def cmd_import_ris(args: argparse.Namespace) -> int:
    imported_items = import_ris_file_in_batches(
        Path(args.ris_file),
        batch_size=int(args.batch_size),
    )
    print(json.dumps(parse_import_results(imported_items), ensure_ascii=False, indent=2))
    return 0


def cmd_run(args: argparse.Namespace) -> int:
    spec = load_json(Path(args.json_file))
    if not isinstance(spec, dict):
        raise RuntimeError("Workflow JSON must be an object")
    result = run_workflow(spec)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    import_parser = subparsers.add_parser(
        "import-records",
        help="Import literature records into Zotero and print the created item keys",
    )
    import_parser.add_argument("--json-file", required=True, help="Path to a JSON array of literature records")
    import_parser.add_argument("--tag", action="append", help="Global tag to apply to all imported records")
    import_parser.set_defaults(func=cmd_import_records)

    import_ris_parser = subparsers.add_parser(
        "import-ris",
        help="Import a RIS file into Zotero in small batches for more reliable local imports",
    )
    import_ris_parser.add_argument("--ris-file", required=True, help="Path to a RIS file")
    import_ris_parser.add_argument(
        "--batch-size",
        type=int,
        default=5,
        help="Number of RIS records to import per connector request (default: 5)",
    )
    import_ris_parser.set_defaults(func=cmd_import_ris)

    run_parser = subparsers.add_parser(
        "run",
        help="Import records, draft a Word summary, insert real Zotero citations, and add a bibliography",
    )
    run_parser.add_argument("--json-file", required=True, help="Path to a workflow JSON object")
    run_parser.set_defaults(func=cmd_run)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return int(args.func(args))
    except RuntimeError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
