#!/usr/bin/env python3
"""Drive real Zotero Word citations via a local Zotero bridge plugin.

This helper installs a Zotero 7 plugin into the active Zotero profile, makes
sure Zotero is running with that plugin loaded, and then talks to the plugin's
local HTTP bridge to insert real Zotero fields into the active Microsoft Word
document on macOS.
"""

from __future__ import annotations

import argparse
import configparser
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any


ADDON_ID = "codex-word-bridge@local.codex"
DEFAULT_STYLE = "APA Style 7th edition"
DEFAULT_PORT = 23119
PORT_CANDIDATES = [DEFAULT_PORT, 23120, 23121, 23122, 23123]
PING_PATH = "/codex/zotero-word/ping"
STATUS_PATH = "/codex/zotero-word/status"
ADDONS_PATH = "/codex/zotero-word/addons"
ENABLE_ADDONS_PATH = "/codex/zotero-word/addons/enable"
INSERT_PATH = "/codex/zotero-word/insert"
STYLE_PATH = "/codex/zotero-word/style"
REFRESH_PATH = "/codex/zotero-word/refresh"
BIBLIOGRAPHY_PATH = "/codex/zotero-word/bibliography"
SERVER_INTEGRATION_ENTRY = "chrome/content/zotero/xpcom/server/server_integration.js"
BRIDGE_BEGIN_MARKER = "// CODEx_ZOTERO_WORD_BRIDGE_BEGIN"
BRIDGE_END_MARKER = "// CODEx_ZOTERO_WORD_BRIDGE_END"
DEFAULT_HTTP_TIMEOUT_S = 60.0
DEFAULT_ZOTERO_APP_PATH = Path("/Applications/Zotero.app")
DEFAULT_WORD_DOCUMENT_ID = "/Applications/Microsoft Word.app/"
WORD_DOCUMENT_ID_PLACEHOLDER = "__CODEX_WORD_DOCUMENT_ID__"
BRIDGE_TEMPLATE_PATHS = {
    "chrome/content/scripts/bridge.js",
    "server_integration_bridge.js",
}


class BridgeError(RuntimeError):
    pass


@dataclass
class BridgeContext:
    profile_dir: Path
    extensions_dir: Path
    addon_path: Path
    base_url: str
    install_mode: str


def env_bool(name: str, default: bool = False) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


def zotero_app_path() -> Path:
    raw = os.environ.get("ZOTERO_APP_PATH")
    if raw:
        return Path(raw).expanduser()
    return DEFAULT_ZOTERO_APP_PATH


def omni_ja_path() -> Path:
    raw = os.environ.get("ZOTERO_OMNI_JA_PATH")
    if raw:
        return Path(raw).expanduser()
    return zotero_app_path() / "Contents/Resources/app/omni.ja"


def omni_backup_path() -> Path:
    return omni_ja_path().with_suffix(".ja.codex-backup")


def word_document_id() -> str:
    value = (
        os.environ.get("WORD_DOCUMENT_ID")
        or os.environ.get("WORD_APP_PATH")
        or DEFAULT_WORD_DOCUMENT_ID
    )
    value = value.strip()
    if not value.endswith("/"):
        value += "/"
    return value


def render_bridge_text(text: str) -> str:
    return text.replace(WORD_DOCUMENT_ID_PLACEHOLDER, json.dumps(word_document_id()))


def skill_dir() -> Path:
    return Path(__file__).resolve().parents[1]


def addon_source_dir() -> Path:
    return skill_dir() / "assets" / "zotero-codex-word-bridge"


def server_integration_bridge_path() -> Path:
    return addon_source_dir() / "server_integration_bridge.js"


def zotero_support_dir() -> Path:
    raw = os.environ.get("ZOTERO_SUPPORT_DIR")
    if raw:
        return Path(raw).expanduser()
    return Path.home() / "Library" / "Application Support" / "Zotero"


def profiles_ini_path() -> Path:
    return zotero_support_dir() / "profiles.ini"


def run_command(args: list[str], *, capture_output: bool = True, check: bool = True) -> subprocess.CompletedProcess[str]:
    proc = subprocess.run(
        args,
        text=True,
        capture_output=capture_output,
    )
    if check and proc.returncode != 0:
        stderr = proc.stderr.strip() if proc.stderr else ""
        raise BridgeError(stderr or f"Command failed: {' '.join(args)}")
    return proc


def run_osascript(script: str) -> str:
    proc = subprocess.run(
        ["osascript", "-"],
        input=script,
        text=True,
        capture_output=True,
    )
    if proc.returncode != 0:
        raise BridgeError(proc.stderr.strip() or "osascript failed")
    return proc.stdout.strip()


def is_app_running(app_name: str) -> bool:
    script = f"""
tell application "System Events"
  return (name of every process) contains "{app_name}"
end tell
"""
    return run_osascript(script).lower() == "true"


def open_app(app_name: str) -> None:
    run_command(["open", "-a", app_name], capture_output=True)


def quit_app(app_name: str) -> None:
    script = f'tell application "{app_name}" to quit'
    subprocess.run(
        ["osascript", "-e", script],
        text=True,
        capture_output=True,
    )


def wait_for_app(app_name: str, *, running: bool, timeout_s: float) -> bool:
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        if is_app_running(app_name) == running:
            return True
        time.sleep(0.25)
    return False


def locate_profile_dir() -> Path:
    ini_path = profiles_ini_path()
    if not ini_path.exists():
        raise BridgeError(f"Zotero profiles.ini not found: {ini_path}")

    parser = configparser.ConfigParser()
    parser.read(ini_path)

    default_path: Path | None = None
    for section in parser.sections():
        if not section.startswith("Profile"):
            continue
        rel = parser.getboolean(section, "IsRelative", fallback=True)
        raw_path = parser.get(section, "Path", fallback="")
        if not raw_path:
            continue
        candidate = zotero_support_dir() / raw_path if rel else Path(raw_path)
        if parser.getboolean(section, "Default", fallback=False):
            default_path = candidate
            break
        if default_path is None:
            default_path = candidate

    if not default_path or not default_path.exists():
        raise BridgeError("Could not locate the active Zotero profile directory")
    return default_path


def addon_xpi_bytes() -> bytes:
    source = addon_source_dir()
    if not source.exists():
        raise BridgeError(f"Bridge addon source not found: {source}")

    with tempfile.NamedTemporaryFile(suffix=".xpi", delete=False) as tmp:
        tmp_path = Path(tmp.name)

    try:
        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for path in sorted(source.rglob("*")):
                if path.is_dir():
                    continue
                rel_path = path.relative_to(source).as_posix()
                if rel_path in BRIDGE_TEMPLATE_PATHS:
                    data = render_bridge_text(path.read_text()).encode("utf-8")
                    zf.writestr(rel_path, data)
                    continue
                zf.write(path, rel_path)
        return tmp_path.read_bytes()
    finally:
        tmp_path.unlink(missing_ok=True)


def addon_target_path(profile_dir: Path) -> Path:
    return profile_dir / "extensions" / f"{ADDON_ID}.xpi"


def install_addon(profile_dir: Path) -> tuple[Path, bool]:
    extensions_dir = profile_dir / "extensions"
    extensions_dir.mkdir(parents=True, exist_ok=True)
    addon_path = addon_target_path(profile_dir)
    new_bytes = addon_xpi_bytes()

    changed = True
    if addon_path.exists():
        changed = addon_path.read_bytes() != new_bytes

    if changed:
        addon_path.write_bytes(new_bytes)

    return addon_path, changed


def read_port_candidates(profile_dir: Path) -> list[int]:
    prefs_path = profile_dir / "prefs.js"
    candidates = []
    if prefs_path.exists():
        for line in prefs_path.read_text(errors="ignore").splitlines():
            if 'httpServer.port' not in line:
                continue
            digits = "".join(ch for ch in line if ch.isdigit())
            if digits:
                candidates.append(int(digits))
    for port in PORT_CANDIDATES:
        if port not in candidates:
            candidates.append(port)
    return candidates


def read_pref_value(profile_dir: Path, pref_name: str) -> str | bool | None:
    prefs_path = profile_dir / "prefs.js"
    if not prefs_path.exists():
        return None

    string_pattern = re.compile(
        rf'user_pref\("{re.escape(pref_name)}",\s*"((?:[^"\\\\]|\\\\.)*)"\s*\);'
    )
    bool_pattern = re.compile(
        rf'user_pref\("{re.escape(pref_name)}",\s*(true|false)\s*\);'
    )

    for line in prefs_path.read_text(errors="ignore").splitlines():
        if pref_name not in line:
            continue
        if match := string_pattern.search(line):
            raw = match.group(1)
            return bytes(raw, "utf-8").decode("unicode_escape")
        if match := bool_pattern.search(line):
            return match.group(1) == "true"
    return None


def zotero_data_dir(profile_dir: Path) -> Path:
    use_data_dir = read_pref_value(profile_dir, "extensions.zotero.useDataDir")
    raw_data_dir = read_pref_value(profile_dir, "extensions.zotero.dataDir")

    if use_data_dir and isinstance(raw_data_dir, str) and raw_data_dir.strip():
        data_dir = Path(raw_data_dir).expanduser()
        if data_dir.exists():
            return data_dir
        raise BridgeError(f"Configured Zotero data directory not found: {data_dir}")

    return profile_dir


def http_json(
    method: str,
    url: str,
    payload: dict[str, Any] | None = None,
    *,
    timeout_s: float = DEFAULT_HTTP_TIMEOUT_S,
) -> dict[str, Any]:
    data = None
    headers = {}
    if payload is not None:
        data = json.dumps(payload).encode("utf-8")
        headers["Content-Type"] = "application/json"

    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=timeout_s) as resp:
            body = resp.read().decode("utf-8")
            if not body:
                return {}
            return json.loads(body)
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise BridgeError(
            f"HTTP {exc.code} calling {url}: {body or exc.reason}"
        ) from exc


def probe_bridge_base(profile_dir: Path) -> str | None:
    for port in read_port_candidates(profile_dir):
        base = f"http://127.0.0.1:{port}"
        try:
            payload = http_json("GET", f"{base}{PING_PATH}", timeout_s=1.5)
        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, json.JSONDecodeError):
            continue
        if payload.get("ok"):
            return base
    return None


def restart_zotero_if_needed(*, force_restart: bool) -> None:
    was_running = is_app_running("Zotero")
    if was_running and force_restart:
        quit_app("Zotero")
        wait_for_app("Zotero", running=False, timeout_s=20.0)
    if force_restart or not was_running:
        open_app("Zotero")


def wait_for_bridge(profile_dir: Path, timeout_s: float = 45.0) -> str:
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        base = probe_bridge_base(profile_dir)
        if base:
            return base
        time.sleep(0.5)
    raise BridgeError("Timed out waiting for the Zotero bridge endpoint to come online")


def bridge_patch_text() -> str:
    patch_path = server_integration_bridge_path()
    if not patch_path.exists():
        raise BridgeError(f"Server integration bridge patch not found: {patch_path}")
    return render_bridge_text(patch_path.read_text())


def upsert_bridge_patch(original_text: str, patch_text: str) -> str:
    if BRIDGE_BEGIN_MARKER in original_text and BRIDGE_END_MARKER in original_text:
        start = original_text.index(BRIDGE_BEGIN_MARKER)
        end = original_text.index(BRIDGE_END_MARKER, start) + len(BRIDGE_END_MARKER)
        before = original_text[:start].rstrip()
        after = original_text[end:].lstrip("\n")
        parts = [before, patch_text.rstrip()]
        if after:
            parts.append(after)
        return "\n\n".join(part for part in parts if part) + "\n"

    return original_text.rstrip() + "\n\n" + patch_text.rstrip() + "\n"


def patch_omni_ja() -> bool:
    active_omni_ja_path = omni_ja_path()
    active_backup_path = omni_backup_path()
    if not active_omni_ja_path.exists():
        raise BridgeError(f"Zotero omni.ja not found: {active_omni_ja_path}")

    patch_text = bridge_patch_text()

    with zipfile.ZipFile(active_omni_ja_path, "r") as source_zip:
        try:
            original_text = source_zip.read(SERVER_INTEGRATION_ENTRY).decode("utf-8")
        except KeyError as exc:
            raise BridgeError(f"Missing {SERVER_INTEGRATION_ENTRY} in {active_omni_ja_path}") from exc

        new_text = upsert_bridge_patch(original_text, patch_text)
        if new_text == original_text:
            return False

        if not active_backup_path.exists():
            shutil.copy2(active_omni_ja_path, active_backup_path)

        with tempfile.NamedTemporaryFile(
            suffix=".omni.ja",
            delete=False,
            dir=str(active_omni_ja_path.parent),
        ) as tmp:
            tmp_path = Path(tmp.name)

        try:
            with zipfile.ZipFile(tmp_path, "w") as dest_zip:
                for info in source_zip.infolist():
                    data = source_zip.read(info.filename)
                    if info.filename == SERVER_INTEGRATION_ENTRY:
                        data = new_text.encode("utf-8")
                    dest_zip.writestr(info, data)
            tmp_path.replace(active_omni_ja_path)
        finally:
            tmp_path.unlink(missing_ok=True)

    return True


def omni_patch_installed() -> bool:
    active_omni_ja_path = omni_ja_path()
    if not active_omni_ja_path.exists():
        return False
    try:
        with zipfile.ZipFile(active_omni_ja_path, "r") as source_zip:
            text = source_zip.read(SERVER_INTEGRATION_ENTRY).decode("utf-8")
    except Exception:
        return False
    return BRIDGE_BEGIN_MARKER in text and BRIDGE_END_MARKER in text


def ensure_bridge_ready(*, force_restart: bool = False, allow_omni_patch: bool | None = None) -> BridgeContext:
    if allow_omni_patch is None:
        allow_omni_patch = env_bool("ZOTERO_WORD_ALLOW_OMNI_JA_PATCH", default=False)

    profile_dir = locate_profile_dir()
    addon_path = addon_target_path(profile_dir)
    base = probe_bridge_base(profile_dir)
    install_mode = "omni-patch" if omni_patch_installed() else "addon"

    # When the bridge is already available via the omni.ja patch, do not touch
    # the addon XPI or restart Zotero for routine requests. This keeps batch
    # citation insertion fast and avoids restarting Zotero between inserts.
    if base and install_mode == "omni-patch" and not force_restart:
        return BridgeContext(
            profile_dir=profile_dir,
            extensions_dir=profile_dir / "extensions",
            addon_path=addon_path,
            base_url=base,
            install_mode=install_mode,
        )

    addon_changed = False
    if install_mode != "omni-patch" or force_restart:
        addon_path, addon_changed = install_addon(profile_dir)

    need_restart = force_restart or addon_changed or base is None
    if need_restart:
        restart_zotero_if_needed(force_restart=need_restart and is_app_running("Zotero"))
        try:
            base = wait_for_bridge(profile_dir, timeout_s=15.0)
        except BridgeError:
            base = None

    if not base:
        if not allow_omni_patch:
            raise BridgeError(
                "Bridge add-on did not come online. Re-run with "
                "--allow-omni-ja-patch or set ZOTERO_WORD_ALLOW_OMNI_JA_PATCH=1 "
                "to enable the omni.ja fallback."
            )
        if is_app_running("Zotero"):
            quit_app("Zotero")
            wait_for_app("Zotero", running=False, timeout_s=20.0)
        patch_omni_ja()
        install_mode = "omni-patch"
        open_app("Zotero")
        base = wait_for_bridge(profile_dir, timeout_s=30.0)

    if not base:
        raise BridgeError("Could not find a running Zotero bridge endpoint")

    return BridgeContext(
        profile_dir=profile_dir,
        extensions_dir=profile_dir / "extensions",
        addon_path=addon_path,
        base_url=base,
        install_mode=install_mode,
    )


def bridge_post(ctx: BridgeContext, path: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
    result = http_json("POST", f"{ctx.base_url}{path}", payload or {})
    if not result.get("ok"):
        raise BridgeError(result.get("error", f"Bridge call failed: {path}"))
    return result


def bridge_get(ctx: BridgeContext, path: str) -> dict[str, Any]:
    result = http_json("GET", f"{ctx.base_url}{path}")
    if not result.get("ok"):
        raise BridgeError(result.get("error", f"Bridge call failed: {path}"))
    return result


def zotero_db_path(profile_dir: Path) -> Path:
    db = zotero_data_dir(profile_dir) / "zotero.sqlite"
    if not db.exists():
        raise BridgeError(f"Zotero database not found: {db}")
    return db


def resolve_item_keys(profile_dir: Path, keys: list[str], library_id: int | None) -> list[int]:
    if not keys:
        return []

    uri = f"file:{zotero_db_path(profile_dir)}?mode=ro&immutable=1"
    conn = sqlite3.connect(uri, uri=True)
    try:
        item_ids = []
        for key in keys:
            if library_id is None:
                rows = conn.execute(
                    "SELECT itemID, libraryID FROM items WHERE key = ?",
                    (key,),
                ).fetchall()
            else:
                rows = conn.execute(
                    "SELECT itemID, libraryID FROM items WHERE key = ? AND libraryID = ?",
                    (key, library_id),
                ).fetchall()

            if not rows:
                raise BridgeError(f"Zotero item key not found: {key}")
            if len(rows) > 1:
                ids = ", ".join(str(row[1]) for row in rows)
                raise BridgeError(
                    f"Zotero item key '{key}' exists in multiple libraries ({ids}); pass --library-id"
                )
            item_ids.append(int(rows[0][0]))
        return item_ids
    finally:
        conn.close()


def collect_item_ids(args: argparse.Namespace, profile_dir: Path) -> list[int]:
    item_ids = list(args.item_ids or [])
    item_ids.extend(resolve_item_keys(profile_dir, list(args.item_keys or []), args.library_id))
    if not item_ids:
        raise BridgeError("Provide at least one --item-id or --item-key")
    return item_ids


def json_dump(payload: dict[str, Any]) -> None:
    print(json.dumps(payload, ensure_ascii=False, indent=2))


def add_omni_patch_flag(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--allow-omni-ja-patch",
        action="store_true",
        help="Allow a fallback patch to Zotero's omni.ja when the add-on path does not load",
    )


def cmd_install_bridge(args: argparse.Namespace) -> int:
    ctx = ensure_bridge_ready(
        force_restart=args.restart,
        allow_omni_patch=args.allow_omni_ja_patch,
    )
    json_dump(
        {
            "ok": True,
            "profile_dir": str(ctx.profile_dir),
            "addon_path": str(ctx.addon_path),
            "base_url": ctx.base_url,
            "install_mode": ctx.install_mode,
        }
    )
    return 0


def cmd_check(args: argparse.Namespace) -> int:
    ctx = ensure_bridge_ready(allow_omni_patch=args.allow_omni_ja_patch)
    payload = {
        "ping": bridge_get(ctx, PING_PATH),
    }
    try:
        payload["status"] = bridge_get(ctx, STATUS_PATH)
    except Exception as exc:
        payload["status_error"] = str(exc)
    payload["profile_dir"] = str(ctx.profile_dir)
    payload["addon_path"] = str(ctx.addon_path)
    payload["install_mode"] = ctx.install_mode
    json_dump(payload)
    return 0


def cmd_status(args: argparse.Namespace) -> int:
    ctx = ensure_bridge_ready(allow_omni_patch=args.allow_omni_ja_patch)
    json_dump(bridge_get(ctx, STATUS_PATH))
    return 0


def cmd_addons(args: argparse.Namespace) -> int:
    ctx = ensure_bridge_ready(allow_omni_patch=args.allow_omni_ja_patch)
    json_dump(bridge_get(ctx, ADDONS_PATH))
    return 0


def cmd_enable_addons(args: argparse.Namespace) -> int:
    ctx = ensure_bridge_ready(allow_omni_patch=args.allow_omni_ja_patch)
    if not args.all and not args.ids:
        raise BridgeError("Pass --all or at least one --id")
    payload = {
        "all": bool(args.all),
        "ids": list(args.ids or []),
    }
    json_dump(bridge_post(ctx, ENABLE_ADDONS_PATH, payload))
    return 0


def cmd_style(args: argparse.Namespace) -> int:
    ctx = ensure_bridge_ready(allow_omni_patch=args.allow_omni_ja_patch)
    json_dump(bridge_post(ctx, STYLE_PATH, {"style": args.style}))
    return 0


def cmd_refresh(args: argparse.Namespace) -> int:
    ctx = ensure_bridge_ready(allow_omni_patch=args.allow_omni_ja_patch)
    json_dump(bridge_post(ctx, REFRESH_PATH, {}))
    return 0


def cmd_bibliography(args: argparse.Namespace) -> int:
    ctx = ensure_bridge_ready(allow_omni_patch=args.allow_omni_ja_patch)
    json_dump(bridge_post(ctx, BIBLIOGRAPHY_PATH, {"style": args.style}))
    return 0


def cmd_insert(args: argparse.Namespace) -> int:
    ctx = ensure_bridge_ready(allow_omni_patch=args.allow_omni_ja_patch)
    item_ids = collect_item_ids(args, ctx.profile_dir)

    if len(item_ids) > 1 and any(
        value is not None
        for value in [args.locator, args.label, args.prefix, args.suffix]
    ):
        raise BridgeError("Locator/prefix/suffix options only support a single cited item")

    citation_items: list[dict[str, Any]] = []
    for index, item_id in enumerate(item_ids):
        citation_item: dict[str, Any] = {"id": item_id}
        if index == 0:
            if args.locator:
                citation_item["locator"] = args.locator
            if args.label:
                citation_item["label"] = args.label
            if args.prefix:
                citation_item["prefix"] = args.prefix
            if args.suffix:
                citation_item["suffix"] = args.suffix
            if args.suppress_author:
                citation_item["suppress-author"] = True
            if args.author_only:
                citation_item["author-only"] = True
        citation_items.append(citation_item)

    payload = {
        "style": args.style,
        "citationItems": citation_items,
    }
    json_dump(bridge_post(ctx, INSERT_PATH, payload))
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    install_parser = subparsers.add_parser("install-bridge", help="Install or update the Zotero bridge plugin")
    install_parser.add_argument("--restart", action="store_true", help="Force a Zotero restart even if the plugin is unchanged")
    add_omni_patch_flag(install_parser)
    install_parser.set_defaults(func=cmd_install_bridge)

    check_parser = subparsers.add_parser("check", help="Verify that Zotero bridge is running and inspect the active Word document")
    add_omni_patch_flag(check_parser)
    check_parser.set_defaults(func=cmd_check)

    status_parser = subparsers.add_parser("status", help="Query bridge status for the active Word document")
    add_omni_patch_flag(status_parser)
    status_parser.set_defaults(func=cmd_status)

    addons_parser = subparsers.add_parser("addons", help="List Zotero add-ons currently recognized by the running Zotero instance")
    add_omni_patch_flag(addons_parser)
    addons_parser.set_defaults(func=cmd_addons)

    enable_addons_parser = subparsers.add_parser("enable-addons", help="Enable Zotero add-ons currently recognized by the running Zotero instance")
    enable_addons_parser.add_argument("--id", dest="ids", action="append", help="Add-on id to enable; repeat for multiple add-ons")
    enable_addons_parser.add_argument("--all", action="store_true", help="Enable all recognized Zotero add-ons")
    add_omni_patch_flag(enable_addons_parser)
    enable_addons_parser.set_defaults(func=cmd_enable_addons)

    style_parser = subparsers.add_parser("ensure-style", help="Set document Zotero style and refresh existing fields")
    style_parser.add_argument("--style", default=DEFAULT_STYLE, help="Citation style name or style ID")
    add_omni_patch_flag(style_parser)
    style_parser.set_defaults(func=cmd_style)

    refresh_parser = subparsers.add_parser("refresh", help="Refresh citations and bibliography in the active Word document")
    add_omni_patch_flag(refresh_parser)
    refresh_parser.set_defaults(func=cmd_refresh)

    bibliography_parser = subparsers.add_parser("bibliography", help="Insert or refresh a Zotero bibliography in the active Word document")
    bibliography_parser.add_argument("--style", default=DEFAULT_STYLE, help="Citation style name or style ID")
    add_omni_patch_flag(bibliography_parser)
    bibliography_parser.set_defaults(func=cmd_bibliography)

    insert_parser = subparsers.add_parser("insert", help="Insert a real Zotero citation into the active Word document")
    insert_parser.add_argument("--item-id", dest="item_ids", type=int, action="append", help="Zotero item ID to cite; repeat for multiple items")
    insert_parser.add_argument("--item-key", dest="item_keys", action="append", help="Zotero item key to cite; repeat for multiple items")
    insert_parser.add_argument("--library-id", type=int, help="Library ID for resolving --item-key when needed")
    insert_parser.add_argument("--style", default=DEFAULT_STYLE, help="Citation style name or style ID")
    insert_parser.add_argument("--locator", help="Locator for a single cited item, such as 123-125")
    insert_parser.add_argument("--label", help="Locator label, such as page or chapter")
    insert_parser.add_argument("--prefix", help="Prefix text for a single cited item")
    insert_parser.add_argument("--suffix", help="Suffix text for a single cited item")
    insert_parser.add_argument("--suppress-author", action="store_true", help="Suppress author in a single cited item")
    insert_parser.add_argument("--author-only", action="store_true", help="Insert author only for a single cited item")
    add_omni_patch_flag(insert_parser)
    insert_parser.set_defaults(func=cmd_insert)

    init_config_parser = subparsers.add_parser("init-config", help="Backward-compatible alias for install-bridge")
    init_config_parser.add_argument("--restart", action="store_true", help="Force a Zotero restart even if the plugin is unchanged")
    add_omni_patch_flag(init_config_parser)
    init_config_parser.set_defaults(func=cmd_install_bridge)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return int(args.func(args))
    except BridgeError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
