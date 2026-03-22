# Setup

## Prerequisites

- Run on macOS.
- Install Python 3.
- Keep Microsoft Word open with the target document visible.
- Keep Zotero open.
- Install the Zotero Word plugin already bundled by Zotero.
- Grant macOS automation permission when Zotero asks to control Microsoft Word.
- Install `zotero-mcp-server`, or set `ZOTERO_MCP_PYTHON` to a Python runtime that already has `zotero_mcp` installed.

## Optional Environment Variables

- `ZOTERO_MCP_PYTHON`: Python executable that can `import zotero_mcp`.
- `ZOTERO_SUPPORT_DIR`: Override the default Zotero support directory if `~/Library/Application Support/Zotero` is not correct.
- `ZOTERO_APP_PATH`: Override the Zotero app bundle path when Zotero is not installed at `/Applications/Zotero.app`.
- `ZOTERO_OMNI_JA_PATH`: Override the exact `omni.ja` path if you need the fallback patch against a non-standard install.
- `WORD_APP_PATH` or `WORD_DOCUMENT_ID`: Override the Microsoft Word bundle path used by the bridge for non-standard installs.
- `ZOTERO_WORD_ALLOW_OMNI_JA_PATCH=1`: Allow the fallback that patches Zotero's `omni.ja`. Leave this unset unless you explicitly want that behavior.

## One-Time Bridge Setup

Install or refresh the local Zotero bridge:

```bash
python3 scripts/zotero_word_plugin.py install-bridge --restart
```

The helper first tries the bundled Zotero add-on path.

If that still does not load on the current machine, the public release does not patch Zotero's app bundle automatically. Review the implications first, then opt in explicitly:

```bash
python3 scripts/zotero_word_plugin.py install-bridge --restart --allow-omni-ja-patch
```

The `omni.ja` fallback modifies the local Zotero application bundle. That is why it is opt-in in this public version.

## Default Commands

```bash
python3 scripts/zotero_word_plugin.py check
python3 scripts/zotero_word_plugin.py ensure-style --style "APA Style 7th edition"
python3 scripts/zotero_word_plugin.py insert --item-id 20394
python3 scripts/zotero_word_plugin.py insert --item-key PNQIAINP
python3 scripts/zotero_word_plugin.py bibliography
python3 scripts/zotero_word_plugin.py refresh
python3 scripts/zotero_word_workflow.py import-ris --ris-file /path/to/export.ris --batch-size 5
```

## Style Default

Default the first Zotero citation style in a document to `APA Style 7th edition`.

- For initial insertion, `insert` will try to handle the first `Document Preferences` popup automatically.
- If the document style must be forced or corrected explicitly, run `ensure-style`.

## Item Resolution Strategy

Prefer:

1. exact Zotero item key
2. exact Zotero item ID
3. exact title
4. `Author Year Exact Title Fragment`

Avoid vague title-only matching when the Zotero library contains many similar items.

The bundled helper can resolve `--item-key` directly against the local Zotero database, including custom Zotero data directories configured through `extensions.zotero.dataDir`.

## Troubleshooting

- If `check` says Word is not ready, open a document window, not just the app shell.
- If the first request triggers a macOS permission dialog, approve Zotero's access to Microsoft Word and rerun the command.
- If APA is not applied on first insert, run `ensure-style --style "APA Style 7th edition"` and retry.
- If the wrong item gets inserted, pass an exact `--item-id` or `--item-key` instead of relying on title text.
- If the bridge endpoint is missing after install, rerun `install-bridge --restart`; on this machine the working path may be the `omni.ja` fallback rather than the XPI addon path.
- If the bridge endpoint is still missing after `install-bridge --restart`, rerun with `--allow-omni-ja-patch` only if you are comfortable letting the helper patch Zotero's app bundle.
- If Zotero uses a custom data directory, the helper reads it from Zotero preferences automatically; verify `extensions.zotero.dataDir` if item-key resolution still fails.
- If the user asks for a real Zotero citation but Word is not open, stop and explain that this skill needs the live Word application.
- If a raw `.ris` import is slow or appears hung, do not keep retrying the whole file in one request. Use `import-ris --batch-size 5`; local Zotero connector imports are much more stable when large exports are split into small batches.
- If the user specifically gave an EBSCO-style RIS export, assume batching is the safer default unless the file is very small.
- If `import-records` fails because the Zotero MCP runtime cannot be found, set `ZOTERO_MCP_PYTHON` to the correct Python executable and retry.
- If Word needs to save an automation-built document to disk, prefer the workflow helper's Save As path instead of `save as active document`. The helper uses UI automation with clipboard paste for the folder path and filename because direct AppleScript `save as` has been unreliable on already-active documents.
- If Save As writes a mangled filename, the automation probably typed the full path into the filename field. Use the helper's directory-plus-basename flow instead of typing one combined path string.
- If the user wants one citation per line, insert a real Return key between entries. Pasting literal `\n` text into Word is not a reliable way to create paragraph breaks.
- If Zotero returns HTTP 500 during insert, rerun with the helper and read the response body. The bridge helper now surfaces the actual HTTP error payload, which is usually enough to tell whether the problem is item resolution, document state, or style initialization.
