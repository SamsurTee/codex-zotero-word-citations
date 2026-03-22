# Zotero Word Citations for Codex

English | [简体中文](README.zh-CN.md)

`zotero-word-citations` is a Codex skill for macOS that works with a live Microsoft Word document, imports records into Zotero, and inserts real Zotero citation fields and bibliographies that remain refreshable through the official Zotero Word integration.

This repository packages the full skill for public installation. It is not affiliated with Zotero or Microsoft.

## What It Does

- Inserts real Zotero citations into the active Word document
- Adds and refreshes bibliographies through Zotero's Word integration
- Imports structured records into Zotero before writing
- Imports RIS files in batches for more reliable local ingestion
- Automates end-to-end workflows that draft a Word document and replace `[[cite:...]]` markers with live Zotero fields

## Who This Is For

This repository is for people who:

- use Codex on macOS
- already work with Zotero and Microsoft Word
- want refreshable Zotero citation fields, not plain-text references
- are comfortable running a few local setup commands

If you only want to edit `.docx` files on disk, this is the wrong tool. This skill operates against live desktop applications.

## Requirements

- macOS
- Microsoft Word desktop
- Zotero desktop with the bundled Zotero Word plugin
- Python 3
- Codex
- `zotero-mcp-server` for `import-records` and `run`

This is not a standalone CLI package. Copying only `SKILL.md` is not enough. The whole repository is the skill.

## Install

### Option 1: Install with the Codex skill installer

```bash
python ~/.codex/skills/.system/skill-installer/scripts/install-skill-from-github.py \
  --repo SamsurTee/codex-zotero-word-citations \
  --path . \
  --name zotero-word-citations
```

Then restart Codex.

### Option 2: Install manually

```bash
git clone https://github.com/SamsurTee/codex-zotero-word-citations.git \
  ~/.codex/skills/zotero-word-citations
```

Then restart Codex.

## First-Time Setup

Run these commands from the repository root:

```bash
python3 scripts/zotero_word_plugin.py install-bridge --restart
python3 scripts/zotero_word_plugin.py check
```

If the bundled add-on route does not come online on your machine, you can explicitly opt in to the fallback that patches Zotero's local `omni.ja`:

```bash
python3 scripts/zotero_word_plugin.py install-bridge --restart --allow-omni-ja-patch
```

That fallback modifies the local Zotero application bundle, so it is opt-in in this public release.

## Quick Start

### Insert a citation into the active Word document

```bash
python3 scripts/zotero_word_plugin.py insert --item-key PNQIAINP
```

### Add or refresh the bibliography

```bash
python3 scripts/zotero_word_plugin.py bibliography
python3 scripts/zotero_word_plugin.py refresh
```

### Import a RIS file in small batches

```bash
python3 scripts/zotero_word_workflow.py import-ris --ris-file /path/to/export.ris --batch-size 5
```

### Run the end-to-end workflow from JSON

```bash
python3 scripts/zotero_word_workflow.py run --json-file /path/to/workflow.json
```

## Typical Codex Prompts

- "Use `$zotero-word-citations` to insert this Zotero item into my open Word document."
- "Find these papers in Zotero, then add real Zotero citations to my draft."
- "Import this RIS file into Zotero in batches, then build a Word summary with refreshable citations."

## Repository Layout

- `SKILL.md`: the skill instructions Codex reads
- `agents/openai.yaml`: Codex UI metadata
- `scripts/`: bridge setup, citation insertion, RIS import, and workflow helpers
- `assets/`: the local Zotero bridge add-on assets
- `references/`: setup and troubleshooting notes

## Environment Overrides

Set these only if your machine differs from the default macOS layout:

- `ZOTERO_MCP_PYTHON`
- `ZOTERO_SUPPORT_DIR`
- `ZOTERO_APP_PATH`
- `ZOTERO_OMNI_JA_PATH`
- `WORD_APP_PATH` or `WORD_DOCUMENT_ID`
- `ZOTERO_WORD_ALLOW_OMNI_JA_PATCH=1`

## Safety Notes

- Keep Word and Zotero open while operating.
- The skill inserts real Zotero fields. It does not synthesize fake citation codes.
- The `omni.ja` fallback changes the local Zotero application bundle. Review that choice before enabling it.

## Troubleshooting

Start with:

```bash
python3 scripts/zotero_word_plugin.py check
```

Then read [references/setup.md](references/setup.md).

## License

This repository is released under `AGPL-3.0-or-later`. See [LICENSE](LICENSE). Zotero's official licensing page is here: [Zotero licensing](https://www.zotero.org/support/licensing).
