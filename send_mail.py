# -*- coding: utf-8 -*-
"""
自动发邮件（引擎上架/交付通知）：
  python send_mail.py init                    # 生成配置模板 mail_config.json
  python send_mail.py test                    # 发一封测试邮件给自己
  python send_mail.py send --subject ... --body-file x.txt [--to a@x,b@x] [--attach f1,f2]
  python send_mail.py engine --lang 泰语 --brand 极氪   # 自动生成引擎上架通知并发送
配置 mail_config.json：
  { "smtp_host": "mail.iflytek.com", "smtp_port": 465, "use_ssl": true,
    "user": "工号@iflytek.com", "password": "邮箱密码或授权码",
    "from": "工号@iflytek.com", "to": ["a@iflytek.com"], "cc": [] }
"""
import os
import sys
import json
import argparse
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.header import Header
from email.utils import formataddr


def app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

APP = app_dir()
if APP not in sys.path:
    sys.path.insert(0, APP)
CONFIG = os.path.join(APP, "mail_config.json")

TEMPLATE = {
    "smtp_host": "mail.iflytek.com",
    "smtp_port": 465,
    "use_ssl": True,
    "user": "你的工号邮箱@iflytek.com",
    "password": "邮箱密码或客户端授权码",
    "from": "你的工号邮箱@iflytek.com",
    "from_name": "多语种识别",
    "to": ["收件人1@iflytek.com"],
    "cc": [],
}


def load_config() -> dict:
    if not os.path.exists(CONFIG):
        return {}
    try:
        with open(CONFIG, encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"✗ 配置读取失败：{e}", flush=True)
        return {}


def init_config() -> int:
    if os.path.exists(CONFIG):
        print(f"配置已存在：{CONFIG}（如需重置请删除后重新 init）", flush=True)
        return 0
    with open(CONFIG, "w", encoding="utf-8") as f:
        json.dump(TEMPLATE, f, ensure_ascii=False, indent=2)
    print(f"✓ 已生成配置模板：{CONFIG}", flush=True)
    print("请打开填好 user/password/to 后再发送。", flush=True)
    return 0


def send(subject: str, body: str, to=None, cc=None, attachments=None,
         html: bool = False) -> int:
    cfg = load_config()
    need = [k for k in ("smtp_host", "user", "password", "from") if not cfg.get(k)]
    if need or "@iflytek.com" not in str(cfg.get("user", "")) and "@" not in str(cfg.get("user", "")):
        pass
    if not cfg or any(not cfg.get(k) or "你的" in str(cfg.get(k, ""))
                      for k in ("user", "password", "from")):
        print(f"✗ 请先配置 {CONFIG}（运行 send_mail.py init 生成模板并填写）。", flush=True)
        return 2
    to = [t for t in (to or cfg.get("to") or []) if t and "@" in t]
    cc = [t for t in (cc or cfg.get("cc") or []) if t and "@" in t]
    if not to:
        print("✗ 没有收件人。", flush=True)
        return 2

    msg = MIMEMultipart()
    msg["Subject"] = Header(subject, "utf-8")
    msg["From"] = formataddr((cfg.get("from_name", ""), cfg["from"]))
    msg["To"] = ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)
    msg.attach(MIMEText(body, "html" if html else "plain", "utf-8"))
    for path in attachments or []:
        if not os.path.exists(path):
            print(f"[warn] 附件不存在，跳过：{path}", flush=True)
            continue
        with open(path, "rb") as f:
            part = MIMEApplication(f.read())
        part.add_header("Content-Disposition", "attachment",
                        filename=Header(os.path.basename(path), "utf-8").encode())
        msg.attach(part)

    host = cfg.get("smtp_host", "mail.iflytek.com")
    port = int(cfg.get("smtp_port", 465))
    print(f"连接 {host}:{port} …", flush=True)
    try:
        if cfg.get("use_ssl", True) and port != 25:
            server = smtplib.SMTP_SSL(host, port, timeout=30)
        else:
            server = smtplib.SMTP(host, port, timeout=30)
            try:
                server.starttls()
            except Exception:
                pass
        server.login(cfg["user"], cfg["password"])
        server.sendmail(cfg["from"], to + cc, msg.as_string())
        server.quit()
    except Exception as e:
        print(f"✗ 发送失败：{e}", flush=True)
        return 1
    print(f"✓ 已发送：{subject}  ->  {', '.join(to)}"
          + (f"（抄送 {', '.join(cc)}）" if cc else ""), flush=True)
    return 0


def engine_mail(lang: str, brand: str, tab: str, extra: str = "",
                to=None, cc=None, dry: bool = False) -> int:
    """按 语种+车厂 取最新引擎，生成上架通知并发送。"""
    import xf_engine
    rows = xf_engine.fetch_sheet(tab)
    recs = [r for r in xf_engine.parse_engines(rows)
            if xf_engine._match(r, lang, brand)]
    if not recs:
        print("✗ 未找到匹配引擎记录。", flush=True)
        return 1
    recs.sort(key=lambda r: r["上架时间"])
    rec = recs[-1]
    subject = f"【引擎上架】{rec['归属']}{rec['语种']} V{rec['V']} R{rec['R']} M{rec['M']}（{rec['上架时间']}）"
    lines = [
        "各位好：",
        "",
        f"{rec['归属']} {rec['语种']} 识别引擎已完成训练并上架，信息如下：",
        "",
        f"  版本：V{rec['V']}  R{rec['R']}  M{rec['M']}",
        f"  上架时间：{rec['上架时间']}",
        f"  更新内容：{rec['描述']}",
        f"  货架位置：{rec['货架']}",
    ]
    if extra:
        lines += ["", f"补充说明：{extra}"]
    lines += ["", "请按需取用，如有问题随时联系。", ""]
    body = "\n".join(lines)
    print("─" * 50, flush=True)
    print("主题:", subject, flush=True)
    print(body, flush=True)
    print("─" * 50, flush=True)
    if dry:
        print("[dry-run] 未发送。", flush=True)
        return 0
    return send(subject, body, to=to, cc=cc)


def main():
    ap = argparse.ArgumentParser(description="发送邮件")
    ap.add_argument("cmd", choices=["init", "test", "send", "engine"])
    ap.add_argument("--subject", default="")
    ap.add_argument("--body", default="")
    ap.add_argument("--body-file", default="")
    ap.add_argument("--to", default="")
    ap.add_argument("--cc", default="")
    ap.add_argument("--attach", default="")
    ap.add_argument("--lang", default="")
    ap.add_argument("--brand", default="")
    ap.add_argument("--tab", default="本地多语种_听写")
    ap.add_argument("--extra", default="")
    ap.add_argument("--dry", action="store_true", help="只预览不发送")
    args = ap.parse_args()

    if args.cmd == "init":
        sys.exit(init_config())

    to = [t.strip() for t in args.to.split(",") if t.strip()] or None
    cc = [t.strip() for t in args.cc.split(",") if t.strip()] or None

    if args.cmd == "test":
        cfg = load_config()
        sys.exit(send("流水线邮件测试", "这是一封测试邮件。", to=to or [cfg.get("user", "")]))

    if args.cmd == "engine":
        sys.exit(engine_mail(args.lang, args.brand, args.tab, args.extra,
                             to=to, cc=cc, dry=args.dry))

    body = args.body
    if args.body_file and os.path.exists(args.body_file):
        body = open(args.body_file, encoding="utf-8").read()
    attach = [a.strip() for a in args.attach.split(",") if a.strip()]
    sys.exit(send(args.subject or "(无主题)", body, to=to, cc=cc, attachments=attach))


if __name__ == "__main__":
    main()
