---
name: yss-zentao-effort-fill
description: Fill Zentao effort/work-hour confirmation from a structured work log. Use when an agent needs to read records with date, time range, task name, and work content, infer or search the target project, fill the Zentao effort confirmation page, click Save, and verify persistence after page reload. Dry-run, no-save, empty data, partial failures, and unverified saves must not be reported as complete.
author: JiyaHe
version: 5.2
last_verified: 2026-07-03
agent_created: true
---

# Zentao Effort Fill Skill

Producer: JiyaHe

## Purpose

Use this skill to fill Zentao effort/work-hour confirmation records from a minimal work log. The expected source columns are:

- `日期` / `date`
- `时间` / `time`
- `任务名称` / `task name`
- `工作内容` / `work content`

The script is self-contained and does not import or depend on any overtime-submission skill.

## Required Configuration

Set organization-specific values before formal use:

```powershell
$env:YSS_ZENTAO_HOST = "zentao.example.com"
$env:YSS_WORKLOG_SOURCE_URL = "<optional-worklog-source-url>"
$env:YSS_ZENTAO_DEFAULT_PROJECT = "<default-project-name>"
$env:YSS_ZENTAO_DEFAULT_PROJECT_QUERY = "<project-search-keyword>"
$env:YSS_ZENTAO_DEFAULT_SYSTEM = "<default-system-name>"
```

Do not hard-code personal document URLs, local paths, credentials, cookies, tokens, customer names, or internal project names in this repository.

## Data Input

Prefer exporting the work log from the agent platform connector to temporary JSON or CSV, then call:

```powershell
$env:PYTHONIOENCODING = "utf-8"
python scripts\fill_zentao_effort.py --start 2026-06-01 --end 2026-06-05 --worklog-json "C:\path\to\worklog.json" --dry-run
```

Formal run:

```powershell
$env:PYTHONIOENCODING = "utf-8"
python scripts\fill_zentao_effort.py --start 2026-06-01 --end 2026-06-05 --worklog-json "C:\path\to\worklog.json"
```

Supported formats:

- Date: `2026年5月28日`, `2026-05-28`, `2026/05/28`, Unix millisecond timestamp.
- Time range: `09:00-12:00`, `09：00-21：00`.
- Hours: subtract lunch `12:00-13:00` and evening `18:00-19:00` by **overlap** with the work period (not only when the full break is covered). Examples: `09:00-18:00` = 8h, `09:00-21:00` = 10h.

## Project Inference

Project inference uses `任务名称` as the primary source. `工作内容` is only description and auxiliary context.

- Tasks matching AI, knowledge-base, digital-employee, wiki, platform, large-model, agent, enablement, or efficiency keywords map to `YSS_ZENTAO_DEFAULT_PROJECT`.
- Tasks containing customer or project keywords use that keyword to search the Zentao project dropdown.
- If no rule matches, fall back to `YSS_ZENTAO_DEFAULT_PROJECT`, and mark it as fallback in dry-run output.
- Never merge all daily work into the fallback project without first classifying each task.

## Browser Flow

The host agent must provide browser automation capability such as Playwright/CDP. The browser should already be logged in, or the agent must pause for manual login and captcha handling.

For each date:

- Open the Zentao effort confirmation page for that date.
- Reuse an existing empty row or click the row-right `+` to add a row.
- Fill name, hours, project, system, task type, and description.
- After selecting project, wait for Zentao to auto-populate system. If missing, read the clickable `对应系统(维护说明)` rule when available and use `YSS_ZENTAO_DEFAULT_SYSTEM`.
- Use row-right `X` to delete surplus default rows before saving.
- Click the page/table Save button. Filling the page is not completion.
- Reload or reopen the same date page and verify persisted rows. If reload readback fails, report failure.

## Strict Completion Rules

Only report completion when all conditions are true:

- At least one valid work-log record was read for the date range.
- Formal mode was used, not dry-run and not `--no-save`.
- Every target date clicked Save.
- The page was reloaded or reopened after saving, and server-side readback validation passed.
- There are no skipped dates, unhandled dialogs, captcha blockers, or unsaved page states.

Otherwise report the exact incomplete state, such as:

- Dry-run only; nothing was filled or saved.
- No fillable work-log records were read.
- A date failed after save/reload validation.
- A date was skipped and needs manual handling or rerun.

## Files

- `scripts/fill_zentao_effort.py`: self-contained parser, inference engine, browser filler, save, and reload validation.
- `工作日志映射.json`: optional runtime-generated task-to-project mapping. Do not commit personal mappings.
