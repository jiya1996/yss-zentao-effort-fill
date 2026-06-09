# yss-zentao-effort-fill

制作人：JiyaHe

用于 WorkBuddy/Codex 的赢时胜禅道工时填写 skill。它读取工作记录中的 `日期 / 时间 / 任务名称 / 工作内容`，按任务名称推理项目，填写禅道工时确认页，并在点击保存后刷新回读校验。

## 使用前注意

- 本 skill 与 `yss-overtime-submission` 完全独立，不要互相 import、复制逻辑或共用状态文件。
- WorkBuddy 使用时应加载本目录的 `SKILL.md`，脚本入口在 `scripts/fill_zentao_effort.py`。
- 正式填写前先 dry-run，确认日期、工时、项目、系统和描述都符合预期。
- 页面填入不等于完成，必须点击“保存”；保存后需要刷新或重新打开同一天页面，回读校验通过才算闭环完成。
- `--no-save` 只用于调试，不能视为任务完成。

## 典型命令

```powershell
$env:PYTHONIOENCODING='utf-8'
python scripts\fill_zentao_effort.py --start 2026-06-01 --end 2026-06-05 --worklog-json "C:\path\to\worklog.json" --dry-run
```

确认计划无误后再去掉 `--dry-run` 正式填写并保存。

## 仓库结构

```text
yss-zentao-effort-fill/
├─ SKILL.md
├─ README.md
├─ .gitignore
└─ scripts/
   └─ fill_zentao_effort.py
```

不要提交 `__pycache__/`、`*.pyc`、浏览器缓存、临时导出的工作日志、账号密码或验证码截图。

## 对外共享前必须脱敏

如果后续要开放给其他同事或放到公开仓库，先检查并删除/替换以下敏感信息：

- 内部系统域名、端口、URL。
- 个人腾讯文档/企微智能表格链接、表格 ID、视图 ID。
- 个人默认本地路径、用户名、账号、密码、token、cookie。
- 公司内部项目名称、客户名称、系统名称、任务描述样例。
- 已提交的历史记录中如果包含敏感信息，需要重写 git history 或新建干净仓库后再共享。

建议对外版本只保留通用字段说明、占位符配置和脱敏样例数据。
