# Zotero Word Citations for Codex

`zotero-word-citations` is a Codex skill for working with a live Microsoft Word document on macOS, importing records into Zotero, and inserting real Zotero citation fields and bibliographies that stay refreshable through the official Zotero Word integration.

This repository is a public packaging of the skill. It is not affiliated with Zotero or Microsoft.

## What This Repository Contains

- `SKILL.md`: the skill instructions Codex reads
- `scripts/`: helper scripts for bridge setup, citation insertion, RIS import, and end-to-end workflows
- `assets/`: the local Zotero bridge add-on assets
- `references/`: setup and troubleshooting notes

## Requirements

- macOS
- Microsoft Word desktop
- Zotero desktop with the bundled Zotero Word plugin
- Python 3
- Codex
- `zotero-mcp-server` for `import-records` and `run`

This is not a standalone command-line package. Copying only `SKILL.md` is not enough. The whole directory is required.

## Install In Codex

Clone or copy this repository into your Codex skills directory as `zotero-word-citations`.

If you prefer to install from GitHub with Codex's skill installer, use the repo root as the skill path:

```bash
python ~/.codex/skills/.system/skill-installer/scripts/install-skill-from-github.py \
  --repo zhentic/codex-zotero-word-citations \
  --path .
```

After installation, restart Codex so it picks up the skill.

## First-Time Setup

From the repository root:

```bash
python3 scripts/zotero_word_plugin.py install-bridge --restart
python3 scripts/zotero_word_plugin.py check
```

If the add-on route does not come online on your machine, you can explicitly allow the fallback that patches Zotero's local `omni.ja`:

```bash
python3 scripts/zotero_word_plugin.py install-bridge --restart --allow-omni-ja-patch
```

That fallback modifies the local Zotero application bundle, so the public release keeps it opt-in.

## Environment Overrides

Set these only if your machine differs from the default macOS layout:

- `ZOTERO_MCP_PYTHON`
- `ZOTERO_SUPPORT_DIR`
- `ZOTERO_APP_PATH`
- `ZOTERO_OMNI_JA_PATH`
- `WORD_APP_PATH` or `WORD_DOCUMENT_ID`
- `ZOTERO_WORD_ALLOW_OMNI_JA_PATCH=1`

## License

This repository is released under `AGPL-3.0-or-later`. Zotero's official licensing page states that Zotero source code is released under the GNU Affero General Public License version 3. See [Zotero licensing](https://www.zotero.org/support/licensing).
