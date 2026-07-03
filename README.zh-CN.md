# yss-zentao-effort-fill

> 根据结构化工作日志，自动填写禅道「工时确认」的便携 Agent Skill。  
> 适用于 WorkBuddy / Codex 类 skill 目录，也可适配 Coze、Copilot、Hermes、Claude Code、OpenClaw、Cursor、Goose 等能读文件、跑 Python、操作浏览器的 Agent 运行时。

[![Agent Skill](https://img.shields.io/badge/Agent%20Skill-portable-green.svg)](https://agentskills.io)

作者：JiyaHe

[English README](README.md)

## 功能概览

- 读取极简工作日志：`日期 / 时间 / 任务名称 / 工作内容`。
- 支持 JSON、CSV、Excel，以及常见智能表格连接器导出的数据。
- 解析中文日期、全角冒号等时间格式。
- 自动计算工时，并扣除午休、晚餐时段（12:00–13:00、18:00–19:00）。
- 根据任务名称推断禅道项目，必要时搜索项目下拉框，并支持可配置的默认项目。
- 通过浏览器自动化填写禅道工时确认页。
- 保存前删除多余的默认行。
- 点击保存后刷新页面，回读校验是否真正落库。
- 对 dry-run、`--no-save`、空数据、部分失败、未校验保存等情况，**不会**误报为已完成。

## 快速开始

### 1. 克隆仓库

```bash
git clone https://github.com/jiya1996/yss-zentao-effort-fill.git
cd yss-zentao-effort-fill
```

### 2. 安装依赖

```bash
pip install playwright openpyxl
playwright install chromium
```

### 3. 配置环境变量

请用环境变量配置，不要把公司内网地址、项目名等私密信息写进源码：

```powershell
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

### 4. 先 dry-run 预览

```bash
PYTHONIOENCODING=utf-8 python scripts/fill_zentao_effort.py \
  --start 2026-06-01 --end 2026-06-05 \
  --worklog-json /path/to/worklog.json \
  --dry-run
```

### 5. 正式填写

确认计划无误后，去掉 `--dry-run`：

```bash
PYTHONIOENCODING=utf-8 python scripts/fill_zentao_effort.py \
  --start 2026-06-01 --end 2026-06-05 \
  --worklog-json /path/to/worklog.json
```

**注意：** 仅把表单填好不算完成。脚本必须点击「保存」，并刷新同一天页面做回读校验。

## 工时计算规则

从 `时间` 列的起止时间计算工时，并按与工作时段的**重叠**扣除休息：

| 休息时段 | 是否计入工时 |
|----------|--------------|
| 12:00–13:00 午休 | 否 |
| 18:00–19:00 晚餐 | 否 |

示例：

| 时间范围 | 计算结果 |
|----------|----------|
| `09:00-18:00` | 8h（扣 1h 午休） |
| `09:00-21:00` | 10h（扣 1h 午休 + 1h 晚餐） |
| `09:00-12:00` + `13:00-18:00` | 3h + 5h = 8h（分段已避开午休，不重复扣） |

## 安装到各 Agent 客户端

### WorkBuddy / Codex

克隆或复制本目录到 Agent 的 skills 目录，然后让 Agent 使用 `yss-zentao-effort-fill`：

```bash
git clone https://github.com/jiya1996/yss-zentao-effort-fill.git ~/.codex/skills/yss-zentao-effort-fill
```

Windows 上的 WorkBuddy，请复制到其本地 skills 目录，并保持文件夹名不变。

### Cursor

将本仓库 clone 到本地 skill 目录，或在对话中引用 `SKILL.md`。需要：

1. 用 Chrome 调试模式启动浏览器（`--remote-debugging-port=9222`），并保持禅道已登录。
2. 先跑 `--dry-run`，再正式执行。

### Coze

Coze 默认不能直接跑本地 Python，可选方案：

- **工具模式**：通过内部 HTTP 服务或工作流节点暴露 `scripts/fill_zentao_effort.py`。
- **浏览器模式**：提供能执行相同 Playwright 步骤的浏览器自动化插件。
- **提示词模式**：把 `SKILL.md` 核心指令贴进 bot，由外部执行器跑脚本。

不要把个人 token、文档链接写进 prompt，应放在平台密钥或环境变量中。

### Copilot / GitHub Copilot Workspace

把本仓库作为项目依赖。Copilot 可读取 `SKILL.md` 并执行：

```bash
python scripts/fill_zentao_effort.py --help
```

实际填表仍需 Playwright 和已登录的浏览器会话。

### Hermes / 其他 Agent 运行时

只要运行时支持以下能力即可适配：

- 加载 `SKILL.md` 作为任务说明；
- 提供工作日志 JSON/CSV/Excel 本地文件；
- 运行 Python；
- 提供浏览器自动化，或连接 Chrome DevTools Protocol（CDP）。

若无法自动化浏览器，可仅用 `--dry-run` 生成计划，再人工填写。

## 工作日志格式

最少需要 4 列：

```csv
日期,时间,任务名称,工作内容
2026-06-01,09:00-12:00,示例任务,示例描述
2026-06-01,13:00-18:00,示例任务,
```

支持的日期/时间示例：

- `2026年5月28日`
- `2026-05-28`
- `09:00-12:00`
- `09：00-21：00`（全角冒号）

## 目录结构

```text
yss-zentao-effort-fill/
├─ SKILL.md              # Agent 技能说明（英文）
├─ README.md             # 英文说明
├─ README.zh-CN.md       # 中文说明（本文件）
├─ .gitignore
└─ scripts/
   └─ fill_zentao_effort.py   # 自包含：解析、推断、填表、保存、回读校验
```

## 安全须知

1. 每次正式运行前，先 `--dry-run`。
2. 保持人工已登录禅道，并随时准备处理验证码、意外弹窗和最终核对。
3. 不要把 `--no-save` 当成「已完成」。
4. 若项目、系统、任务类型、描述明显不对，不要点保存。
5. 禅道升级后 selector 可能失效；请重新 dry-run，并检查浏览器页面结构。

## 敏感信息检查清单

分享给同事或公开发布前，请删除或替换：

- 内网域名、端口、URL；
- 个人智能表格/文档链接、sheet ID、view ID、导出的工作日志；
- 本地绝对路径、用户名、账号、密码、token、cookie、含验证码的截图；
- 内部项目名、客户名、系统名、真实任务描述；
- 运行时生成的 `工作日志映射.json`；
- Git 历史中曾出现过的敏感信息（必要时新建干净仓库，而非只改最新提交）。

本仓库的设计原则是：私密配置通过环境变量和本地临时文件管理，不写入源码。

## 适配到其他公司

优先通过环境变量修改：

| 变量 | 含义 |
|------|------|
| `YSS_ZENTAO_HOST` | 禅道域名（含端口） |
| `YSS_ZENTAO_DEFAULT_PROJECT` | 默认项目名称 |
| `YSS_ZENTAO_DEFAULT_PROJECT_QUERY` | 项目下拉搜索关键词 |
| `YSS_ZENTAO_DEFAULT_SYSTEM` | 默认对应系统 |
| `YSS_WORKLOG_PATH` | 工作日志 Excel 路径 |

若禅道页面结构差异较大，需修改 `scripts/fill_zentao_effort.py` 中的 selector，并用 `--dry-run` 加不保存的浏览器测试验证。

## 许可证

尚未声明许可证。公开发布前请补充 License。
