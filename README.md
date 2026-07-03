# yss-zentao-effort-fill

> A portable Agent Skill for filling Zentao effort/work-hour confirmation from a structured work log.
> Designed for WorkBuddy/Codex-style skills, and adaptable to Coze, Copilot, Hermes, Claude Code, OpenClaw, Cursor, Goose, and other agent runtimes that can read files, run Python, and automate a browser.

[![Agent Skill](https://img.shields.io/badge/Agent%20Skill-portable-green.svg)](https://agentskills.io)

Producer: JiyaHe

[中文说明](README.zh-CN.md)

## What It Does

- Reads minimal work-log records with `日期 / 时间 / 任务名称 / 工作内容`.
- Supports JSON, CSV, Excel, and common smart-sheet connector payloads.
- Parses Chinese dates and full-width time separators.
- Calculates effort hours with lunch/evening break deduction.
- Infers Zentao project from task name, searches project dropdown when needed, and uses a configurable fallback project.
- Fills Zentao effort confirmation rows through browser automation.
- Deletes surplus default rows before saving.
- Clicks Save and reloads the page for persisted readback validation.
- Refuses to claim completion for dry-run, `--no-save`, empty data, partial failures, or unverified saves.

## Quick Start

### 1. Clone

```bash
git clone https://github.com/jiya1996/yss-zentao-effort-fill.git
cd yss-zentao-effort-fill
```

### 2. Install Dependencies

```bash
pip install playwright openpyxl
playwright install chromium
```

### 3. Configure Your Environment

Use environment variables instead of editing private values into source files:

```bash
# Windows PowerShell
$env:YSS_ZENTAO_HOST = "zentao.example.com"
$env:YSS_ZENTAO_DEFAULT_PROJECT = "YOUR_DEFAULT_PROJECT"
$env:YSS_ZENTAO_DEFAULT_PROJECT_QUERY = "YOUR_PROJECT_SEARCH_KEYWORD"
$env:YSS_ZENTAO_DEFAULT_SYSTEM = "YOUR_DEFAULT_SYSTEM"
$env:YSS_WORKLOG_PATH = "C:\path\to\worklog.xlsx"
```

```bash
# macOS / Linux
export YSS_ZENTAO_HOST=zentao.example.com
export YSS_ZENTAO_DEFAULT_PROJECT=YOUR_DEFAULT_PROJECT
export YSS_ZENTAO_DEFAULT_PROJECT_QUERY=YOUR_PROJECT_SEARCH_KEYWORD
export YSS_ZENTAO_DEFAULT_SYSTEM=YOUR_DEFAULT_SYSTEM
export YSS_WORKLOG_PATH=/path/to/worklog.xlsx
```

### 4. Dry Run First

```bash
PYTHONIOENCODING=utf-8 python scripts/fill_zentao_effort.py \
  --start 2026-06-01 --end 2026-06-05 \
  --worklog-json /path/to/worklog.json \
  --dry-run
```

### 5. Formal Run

Only remove `--dry-run` after checking the plan:

```bash
PYTHONIOENCODING=utf-8 python scripts/fill_zentao_effort.py \
  --start 2026-06-01 --end 2026-06-05 \
  --worklog-json /path/to/worklog.json
```

Filling the page is not enough. The script must click Save and reload the same date page for readback validation.

## Install Into Agent Clients

### WorkBuddy / Codex

Clone or copy this folder into the agent skills directory, then ask the agent to use `yss-zentao-effort-fill`.

```bash
git clone https://github.com/jiya1996/yss-zentao-effort-fill.git ~/.codex/skills/yss-zentao-effort-fill
```

For WorkBuddy on Windows, copy or clone into its local skills directory and keep the same folder name.

### Coze

Coze does not run local Python scripts by default. Use one of these integration modes:

- Tool mode: expose `scripts/fill_zentao_effort.py` through an internal HTTP service or workflow node.
- Browser mode: provide a browser automation tool/plugin that can run the same Playwright steps.
- Prompt mode: paste the core instructions from `SKILL.md`, then call an external executor for the script.

The Coze bot should never store personal tokens or document URLs in prompt text. Put them in platform secrets or environment variables.

### Copilot / GitHub Copilot Workspace

Use this repo as a project dependency. Copilot can read `SKILL.md` and run:

```bash
python scripts/fill_zentao_effort.py --help
```

For live browser filling, the runtime still needs Playwright and an authenticated browser session.

### Hermes / Other Agent Runtimes

Any runtime can adapt this skill if it supports:

- Loading `SKILL.md` as task instructions.
- Supplying work-log JSON/CSV/Excel as a local file.
- Running Python.
- Providing browser automation or connecting to a Chrome DevTools Protocol endpoint.

If the runtime cannot automate browsers, use only `--dry-run` and hand the generated plan to a human.

## Expected Work Log

Minimal columns:

```csv
日期,时间,任务名称,工作内容
2026-06-01,09:00-12:00,Example task,Example description
2026-06-01,13:00-18:00,Example task,
```

Supported date/time examples:

- `2026年5月28日`
- `2026-05-28`
- `09:00-12:00`
- `09：00-21：00`

## Repository Structure

```text
yss-zentao-effort-fill/
├─ SKILL.md
├─ README.md
├─ README.zh-CN.md
├─ .gitignore
└─ scripts/
   └─ fill_zentao_effort.py
```

## Safety Notes

1. Do a dry-run before every formal run.
2. Keep a human logged in and available for captcha, unexpected dialogs, and final audit.
3. Never treat `--no-save` as completion.
4. Do not save if project, system, task type, or description looks wrong.
5. UI selectors may drift after Zentao upgrades; rerun dry-run and inspect browser snapshots when that happens.

## Sensitive Information Checklist

Before sharing with colleagues or making a public release, remove or replace:

- Internal hostnames, ports, and URLs.
- Personal smart-sheet/document links, sheet IDs, view IDs, and exported work logs.
- Local absolute paths, usernames, account names, passwords, tokens, cookies, and screenshots containing captcha or session data.
- Internal project names, customer names, system names, and real task descriptions.
- Runtime-generated mapping files such as `工作日志映射.json`.
- Git history that previously contained sensitive values. If needed, create a clean repository instead of only editing the latest commit.

This repository is intended to keep private configuration outside source code through environment variables and temporary local files.

## Adapting To Another Organization

Update configuration through environment variables first:

- `YSS_ZENTAO_HOST`
- `YSS_ZENTAO_DEFAULT_PROJECT`
- `YSS_ZENTAO_DEFAULT_PROJECT_QUERY`
- `YSS_ZENTAO_DEFAULT_SYSTEM`
- `YSS_WORKLOG_PATH`

If your Zentao page differs materially, patch selectors in `scripts/fill_zentao_effort.py` and validate with `--dry-run` plus a non-saving browser test.

## License

No license is declared yet. Add a license before public distribution.
