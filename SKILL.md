---
name: zotero-word-citations
description: Search or import literature into Zotero and insert real Zotero citation fields and bibliographies into a live Microsoft Word document on macOS. Use when the user wants Codex to find papers, import them into Zotero, draft a literature summary in Word, insert refreshable citations through the official Zotero Word plugin, or reformat those citations later. Do not use for plain-text references or generic .docx editing that does not require the live Zotero Word plugin.
---

# Zotero Word Citations

## Overview

Use this skill only for real Zotero Word-plugin citations. For existing library items, search the user's Zotero library via MCP first, resolve the exact Zotero item IDs or keys, and then drive the official Zotero Word integration so the result stays refreshable in Word and can still be reformatted later by Zotero.

For end-to-end natural-language requests such as:
- "检索 5 篇文章，导入 Zotero，然后在 Word 里写综述并插入参考文献"
- "Summarize these papers in Word and cite them with real Zotero fields"
- "Import these DOIs into Zotero and add them to my Word draft"

prefer the high-level workflow entrypoint:

```bash
python3 scripts/zotero_word_workflow.py run --json-file /path/to/workflow.json
```

This wrapper:
- imports structured paper records into Zotero using the Zotero backend
- creates or uses a live Word document
- defaults the first citation style to `APA Style 7th edition`
- inserts real Zotero fields at `[[cite:record-id]]` markers
- inserts a bibliography and refreshes the document
- can save the finished Word document when the workflow JSON includes `output_path`

Run bundled commands from this skill directory, or use the script's absolute path.

## Workflow

1. Run the readiness check:

   ```bash
   python3 scripts/zotero_word_plugin.py check
   ```

   If the check fails, read [references/setup.md](references/setup.md) and fix the environment before proceeding.
   On public installs, review the setup notes for optional environment variables such as `ZOTERO_MCP_PYTHON`, `ZOTERO_SUPPORT_DIR`, `ZOTERO_APP_PATH`, and `WORD_APP_PATH`.

2. Use the Zotero MCP to search the library and identify the exact item.

   If the paper is not in Zotero yet, either:
   - import it with the high-level workflow wrapper, or
   - import structured records first:

   ```bash
   python3 scripts/zotero_word_workflow.py import-records --json-file /path/to/records.json
   ```

   If the user gives you a raw `.ris` export, do not push a large file through Zotero in one request. Prefer the bundled batched importer:

   ```bash
   python3 scripts/zotero_word_workflow.py import-ris --ris-file /path/to/export.ris --batch-size 5
   ```

   Preferred disambiguation order:
   - item key
   - exact title
   - first author + year + distinctive title fragment
   - collection or library context if the MCP exposes them

3. Move the insertion point in the active Word document to the target location.

   This skill inserts at the current Word cursor position. If the user wants citations added to a file on disk without opening Word, stop and explain that this skill requires a live Word window.

4. Insert the citation through the official Word plugin:

   ```bash
   python3 scripts/zotero_word_plugin.py insert --item-id 20394
   python3 scripts/zotero_word_plugin.py insert --item-key PNQIAINP
   ```

   For multiple cited sources:

   ```bash
   python3 scripts/zotero_word_plugin.py insert --item-id 20394 --item-id 19748 --item-id 19708
   ```

   Default the citation style to `APA Style 7th edition` for a document's first Zotero insertion unless the user requests a different style. If Zotero opens `Document Preferences`, the helper script sets APA automatically and then injects the selected Zotero items into the official Quick Format flow without OCR or title search.

   Use `--locator`, `--label`, `--prefix`, `--suffix`, `--suppress-author`, or `--author-only` only for a single cited item.

5. Set or correct the document style explicitly when needed:

   ```bash
   python3 scripts/zotero_word_plugin.py ensure-style --style "APA Style 7th edition"
   ```

6. Add or update the bibliography when requested:

   ```bash
   python3 scripts/zotero_word_plugin.py bibliography
   ```

7. Refresh the document before finishing:

   ```bash
   python3 scripts/zotero_word_plugin.py refresh
   ```

8. If the user wants a file saved to disk, set `output_path` in the workflow JSON or use the workflow helper's Save As path.

   Do not rely on a naive AppleScript `save as active document` step after many separate automation calls. In this environment, the stable path is the workflow helper's UI-driven Save As automation, which pastes the directory and filename separately.

## Rules

- Never synthesize fake Zotero field codes.
- Never replace a requested real Zotero citation with plain text unless the user explicitly agrees.
- Always search Zotero via MCP before inserting, instead of trusting a vague Quick Format search.
- Prefer exact item IDs or item keys over title-driven lookup.
- Default the first document citation style to `APA Style 7th edition` unless the user explicitly requests another style.
- If multiple items plausibly match, stop and ask the user to choose.
- Keep Word and Zotero open while operating.
- Treat this as a live-application workflow, not a static `.docx` XML edit workflow.
- Do not rely on OCR, image recognition, or manual Quick Format typing for normal operation.
- For raw RIS files, prefer `import-ris --batch-size 5` over one-shot connector imports when the export contains many records.
- For one-entry-per-line Word output, use actual Return keystrokes between entries instead of pasting `\n` and assuming Word will keep paragraph boundaries.
- If a bridge request fails with HTTP 500, inspect the returned body; the helper now surfaces the response text instead of only a generic stack trace.

## Commands

```bash
python3 scripts/zotero_word_plugin.py check
python3 scripts/zotero_word_plugin.py init-config
python3 scripts/zotero_word_plugin.py ensure-style --style "APA Style 7th edition"
python3 scripts/zotero_word_plugin.py insert --item-id 20394
python3 scripts/zotero_word_plugin.py insert --item-key PNQIAINP
python3 scripts/zotero_word_plugin.py bibliography
python3 scripts/zotero_word_plugin.py refresh
python3 scripts/zotero_word_workflow.py import-records --json-file /path/to/records.json
python3 scripts/zotero_word_workflow.py import-ris --ris-file /path/to/export.ris --batch-size 5
python3 scripts/zotero_word_workflow.py run --json-file /path/to/workflow.json
```

## Setup And Troubleshooting

Read [references/setup.md](references/setup.md) when:
- `check` fails
- a raw RIS import is unusually slow or appears hung
- the Word document needs to be saved to disk and Save As is behaving inconsistently
- the wrong Zotero item is selected
- macOS automation permissions have not been granted yet
- the bundled add-on path does not load and you need to decide whether to allow the `omni.ja` fallback
