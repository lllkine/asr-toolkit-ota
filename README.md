# ASR 语料工具 · OTA 更新源

本仓库只存放「一键处理语料」GUI 的**热更新脚本**（几百 KB），不含运行时/浏览器/大文件。
老用户点程序里的 **检查更新** 即可自动同步这里的最新脚本，无需重发几百 MB 的整包。

## 用户如何接入

在程序目录（exe 同级）新建 `update_source.txt`，写入本仓库的分支 zip 地址：

```
https://github.com/<你的账号>/<仓库名>/archive/refs/heads/main.zip
```

之后点 **检查更新**：程序会下载该 zip、比对 `version.txt`，有新版就把下列文件覆盖到本地并提示重启。

## 发版流程（维护者）

1. 修改脚本；
2. 抬高 `version.txt`（如 `20260706.1` → `20260707.1`）；
3. `git commit` + `git push`。

用户下次点检查更新即可拿到。

## 文件清单（与程序 `UPDATE_FILES` 一致）

| 文件 | 作用 |
|---|---|
| `pipeline.py` | 语料处理主流水线 |
| `web_download.py` | RMP 语料下载 |
| `qq_schedule.py` | 腾讯文档排期读取 |
| `xf_engine.py` | 引擎查找/下载（登录即下载活会话流程） |
| `send_mail.py` | 引擎上架通知邮件（凭据在本地 `mail_config.json`，不在此仓库） |
| `asr_pipeline_gui.py` | PySide6 图形界面 |
| `tts_tools/asr_tts_tool.py` | TTS 合成工具 |
| `version.txt` | 版本号（OTA 比对依据） |

> 说明：`send_mail.py` 里的邮箱/密码均为占位示例，真实凭据只存在于各用户本地的 `mail_config.json`，不随本仓库分发。

## 内网地址配置（endpoints.json）

脚本里**不含任何内网地址**（RMP / 排期表 / 引擎表 / 制品库 / SMTP 等），统一从程序目录的
`endpoints.json` 读取。本仓库只提供占位模板 `endpoints.example.json`，真实的
`endpoints.json` 由发布方打进 exe、并被 `.gitignore` 挡在仓库之外，OTA 也不会覆盖它
（不在更新清单内）。内网用户拿到的 exe 已内置该文件，开箱即用；二次开发者复制
`endpoints.example.json` 为 `endpoints.json` 填入真实地址即可。
