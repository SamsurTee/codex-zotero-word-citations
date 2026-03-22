# Zotero Word Citations for Codex

[English](README.md) | 简体中文

`zotero-word-citations` 是一个面向 macOS 的 Codex skill，用来操作“正在打开的” Microsoft Word 文档，把文献导入 Zotero，并通过 Zotero 官方 Word 集成插入可刷新、可继续重排格式的真实引用字段和参考文献。

这个仓库是该 skill 的公开发布版本，不是 Zotero 或 Microsoft 的官方项目。

## 它能做什么

- 向当前激活的 Word 文档插入真实的 Zotero 引文
- 通过 Zotero 的 Word 集成插入和刷新参考文献列表
- 先把结构化文献记录导入 Zotero，再继续写作
- 对 RIS 文件做分批导入，提高本地导入稳定性
- 支持从 JSON 工作流生成 Word 草稿，并把 `[[cite:...]]` 标记替换成真实 Zotero 字段

## 适合谁用

这个仓库适合以下用户：

- 在 macOS 上使用 Codex
- 已经在使用 Zotero 和 Microsoft Word
- 需要“可刷新”的 Zotero 引用字段，而不是纯文本参考文献
- 可以接受执行少量本地初始化命令

如果你的需求只是修改磁盘上的 `.docx` 文件，这个 skill 不适合你。它操作的是正在运行的桌面应用。

## 运行要求

- macOS
- Microsoft Word 桌面版
- Zotero 桌面版，并已安装随 Zotero 提供的 Word 插件
- Python 3
- Codex
- `zotero-mcp-server`，用于 `import-records` 和 `run`

这不是一个独立 CLI 工具包。只复制 `SKILL.md` 是不够的，必须安装整个仓库目录。

## 安装方式

### 方式 1：用 Codex skill installer 安装

```bash
python ~/.codex/skills/.system/skill-installer/scripts/install-skill-from-github.py \
  --repo SamsurTee/codex-zotero-word-citations \
  --path . \
  --name zotero-word-citations
```

安装完成后，重启 Codex。

### 方式 2：手动安装

```bash
git clone https://github.com/SamsurTee/codex-zotero-word-citations.git \
  ~/.codex/skills/zotero-word-citations
```

安装完成后，重启 Codex。

## 首次初始化

在仓库根目录运行：

```bash
python3 scripts/zotero_word_plugin.py install-bridge --restart
python3 scripts/zotero_word_plugin.py check
```

如果机器上通过 add-on 的方式始终无法把 bridge 拉起来，你可以显式启用会修改本机 Zotero `omni.ja` 的 fallback：

```bash
python3 scripts/zotero_word_plugin.py install-bridge --restart --allow-omni-ja-patch
```

因为这个 fallback 会修改本地 Zotero 应用包，所以在公开版里它是默认关闭、需要你手动选择开启的。

## 快速开始

### 向当前 Word 文档插入一条引用

```bash
python3 scripts/zotero_word_plugin.py insert --item-key PNQIAINP
```

### 插入或刷新参考文献

```bash
python3 scripts/zotero_word_plugin.py bibliography
python3 scripts/zotero_word_plugin.py refresh
```

### 将 RIS 文件按小批次导入

```bash
python3 scripts/zotero_word_workflow.py import-ris --ris-file /path/to/export.ris --batch-size 5
```

### 通过 JSON 工作流跑完整流程

```bash
python3 scripts/zotero_word_workflow.py run --json-file /path/to/workflow.json
```

## 在 Codex 里可以怎么提

- “用 `$zotero-word-citations` 把这个 Zotero 条目插入到我当前打开的 Word 文档里。”
- “先在 Zotero 里找到这些文献，再把真实 Zotero 引用插入我的草稿。”
- “把这个 RIS 文件分批导入 Zotero，然后生成带可刷新引用的 Word 综述。”

## 仓库结构

- `SKILL.md`：Codex 读取的 skill 指令
- `agents/openai.yaml`：Codex UI 元数据
- `scripts/`：bridge 安装、引用插入、RIS 导入、工作流辅助脚本
- `assets/`：本地 Zotero bridge add-on 资源
- `references/`：安装说明和排错说明

## 可选环境变量

只有在你的机器目录结构和默认 macOS 布局不一致时才需要设置：

- `ZOTERO_MCP_PYTHON`
- `ZOTERO_SUPPORT_DIR`
- `ZOTERO_APP_PATH`
- `ZOTERO_OMNI_JA_PATH`
- `WORD_APP_PATH` 或 `WORD_DOCUMENT_ID`
- `ZOTERO_WORD_ALLOW_OMNI_JA_PATCH=1`

## 安全说明

- 操作期间要保持 Word 和 Zotero 处于打开状态。
- 这个 skill 插入的是真实 Zotero 字段，不会伪造引用代码。
- `omni.ja` fallback 会改动本地 Zotero 应用包，启用前请先确认你接受这个行为。

## 排错建议

先运行：

```bash
python3 scripts/zotero_word_plugin.py check
```

然后查看 [references/setup.md](references/setup.md)。

## License

本仓库采用 `AGPL-3.0-or-later`。见 [LICENSE](LICENSE)。Zotero 官方许可说明见：[Zotero licensing](https://www.zotero.org/support/licensing)。
