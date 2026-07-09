# -*- coding: utf-8 -*-
"""
腾讯文档排期表读取（公开只读表格，无需登录）：
  python qq_schedule.py tabs  --url <腾讯表格链接>            # 列出所有 tab
  python qq_schedule.py read  --url <链接> [--tab 名或id] [--user 责任人]
                              [--scope all|local|cloud] [--download]
read: 解析 单号/需求名称/车厂/语种/云端or本地/预计完成/状态/责任人，
      可按 责任人、云端/本地 过滤；自动把命中行写入排期.xlsx；
      --download 时再从需求管理平台下载对应语料到 _inbox。
"""
import os
import re
import sys
import json
import base64
import zlib
import argparse


def app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


try:  # GBK 控制台下 ✓/↓ 等字符会崩，统一 UTF-8 输出
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

APP = app_dir()
if APP not in sys.path:          # runpy 运行时脚本目录不在 sys.path，补上以便 import web_download
    sys.path.insert(0, APP)
INBOX = os.path.join(APP, "_inbox")
SESSION_DIR = os.path.join(APP, "_session")
STATE = os.path.join(SESSION_DIR, "storage_state.json")


def _endpoint(key, default=""):
    """内网地址：优先环境变量，其次同目录 endpoints.json；公开仓库不含真实地址。"""
    import json as _json
    v = os.environ.get(key, "").strip()
    if v:
        return v
    try:
        _base = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
            else os.path.dirname(os.path.abspath(__file__))
    except Exception:
        _base = os.getcwd()
    for _b in (_base, os.getcwd()):
        try:
            _p = os.path.join(_b, "endpoints.json")
            if os.path.exists(_p):
                _d = _json.load(open(_p, encoding="utf-8"))
                if _d.get(key):
                    return str(_d[key])
        except Exception:
            pass
    return default


DEFAULT_DOC = _endpoint("QQ_SHEET_URL")
RMP_BASE = _endpoint("RMP_BASE")

LANG_SET = {"英语", "葡语", "葡萄牙语", "马来语", "泰语", "阿语", "俄语", "印地语",
            "印尼语", "越南语", "法语", "德语", "西班牙语", "西语", "日语", "韩语",
            "意大利语", "波兰语", "土耳其语", "波斯语", "匈牙利语", "希伯来语",
            "丹麦语", "挪威语", "瑞典语", "荷兰语", "斯洛文尼亚语", "繁体中文", "粤语"}
CLOUD_SET = {"云端", "本地", "云端+本地", "本地+云端"}
STATUS_SET = {"已分配", "未分配", "已完成", "进行中", "处理中", "已解决",
              "已关闭", "已撤销", "验收中", "开发中", "待处理"}
DATE_RE = re.compile(r'^\d{1,2}月\d{1,2}日$|^\d{4}[-/年.]')
SEQ_RE = re.compile(r'^(?:N-CUS|NS)-\d+.*ASR')


# ══════════════════════ protobuf 解码 ══════════════════════

def _read_varint(b, i):
    v = 0; shift = 0
    while i < len(b):
        x = b[i]; i += 1
        v |= (x & 0x7F) << shift
        if not x & 0x80:
            return v, i
        shift += 7
    raise ValueError("bad varint")


def _extract_strings(b, start, end, out, depth=0):
    """按 wire 顺序抽出所有文本（子消息优先递归，失败再当字符串）。"""
    i = start
    while i < end:
        try:
            tag, i = _read_varint(b, i)
        except Exception:
            return False
        wt = tag & 7
        if wt == 0:
            try:
                _, i = _read_varint(b, i)
            except Exception:
                return False
        elif wt == 2:
            try:
                ln, i = _read_varint(b, i)
            except Exception:
                return False
            if ln < 0 or i + ln > end:
                return False
            chunk = b[i:i+ln]
            handled = False
            if depth < 8 and ln > 1:
                sub = []
                # 子消息解析成功且真的产出了字符串，才当作消息；否则回退当文本
                if _extract_strings(b, i, i + ln, sub, depth + 1) and sub:
                    out.extend(sub)
                    handled = True
            if not handled:
                try:
                    s = chunk.decode("utf-8")
                    if s and sum(1 for c in s if c.isprintable()) >= max(1, int(len(s) * 0.9)):
                        out.append(s)
                except Exception:
                    pass
            i += ln
        elif wt == 5:
            i += 4
        elif wt == 1:
            i += 8
        else:
            return False
    return True


def _find_blobs(obj, out):
    """递归找所有 related_sheet 压缩块。"""
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k == "related_sheet" and isinstance(v, str) and len(v) > 100:
                out.append(v)
            else:
                _find_blobs(v, out)
    elif isinstance(obj, list):
        for v in obj:
            _find_blobs(v, out)


def _cells_from_blob(b64: str) -> list:
    dec = zlib.decompress(base64.b64decode(b64))
    out = []
    _extract_strings(dec, 0, len(dec), out)
    # 去掉控制串/纯符号
    return [s.strip() for s in out if s.strip() and any(c.isalnum() for c in s)]


# ══════════════════════ 抓取 ══════════════════════

def fetch_doc(url: str, tab: str = "") -> dict:
    """加载表格页面，截获数据接口。返回 {tabs:[(id,name,hidden)], cells:[str]}"""
    from playwright.sync_api import sync_playwright
    base = url.split("?")[0]
    result = {"tabs": [], "cells": []}
    payloads = {"opendoc": [], "getsheet": []}

    def crawl(page, target_url):
        def on_resp(r):
            u = r.url
            try:
                if "dop-api/opendoc" in u:
                    payloads["opendoc"].append(r.body())
                elif "dop-api/get/sheet" in u:
                    payloads["getsheet"].append((u, r.body()))
            except Exception:
                pass
        page.on("response", on_resp)
        try:
            page.goto(target_url, wait_until="domcontentloaded", timeout=60000)
        except Exception:
            pass
        page.wait_for_timeout(9000)

    with sync_playwright() as p:
        b = p.chromium.launch(headless=True)
        ctx = b.new_context(viewport={"width": 1400, "height": 900})
        page = ctx.new_page()
        crawl(page, base + (f"?tab={tab}" if tab else ""))

        # 解析 tab 列表（opendoc 的 header.d 数组，括号计数提取）
        tab_list = []
        for body in payloads["opendoc"]:
            txt = body.decode("utf-8", errors="replace")
            anchor = txt.find('"header":[{"d":[')
            if anchor < 0:
                continue
            i = txt.find('[', anchor + len('"header":[{"d":'))
            depth, j = 0, i
            while j < len(txt):
                if txt[j] == '[':
                    depth += 1
                elif txt[j] == ']':
                    depth -= 1
                    if depth == 0:
                        break
                j += 1
            try:
                for t in json.loads(txt[i:j+1]):
                    tab_list.append((t.get("id", ""), t.get("name", ""),
                                     bool(t.get("hidden"))))
            except Exception:
                pass
            if tab_list:
                break
        result["tabs"] = tab_list

        # tab 用名字给的 → 解析到 id 后重新加载
        want = tab
        if tab and tab_list and not any(t[0] == tab for t in tab_list):
            hit = next((t for t in tab_list if tab in t[1]), None)
            if hit:
                want = hit[0]
                payloads["getsheet"].clear()
                page2 = ctx.new_page()
                crawl(page2, base + f"?tab={want}")

        # 取数据：解析全部候选数据块，选“单号锚点最多”的那份
        candidates = []   # (是否目标tab, cells)
        for u, body in payloads["getsheet"]:
            txt = body.decode("utf-8", errors="replace")
            m = re.match(r'^[\w$.]+\((.*)\)\s*$', txt, re.S)
            blobs = []
            try:
                _find_blobs(json.loads(m.group(1) if m else txt), blobs)
            except Exception:
                continue
            for blob in blobs:
                try:
                    cells = _cells_from_blob(blob)
                except Exception:
                    continue
                match_tab = (not want) or (f"subId={want}" in u)
                candidates.append((match_tab, cells))
        for body in payloads["opendoc"]:
            txt = body.decode("utf-8", errors="replace")
            m = re.match(r'^[\w$.]+\((.*)\)\s*$', txt, re.S)
            blobs = []
            try:
                _find_blobs(json.loads(m.group(1) if m else txt), blobs)
            except Exception:
                continue
            for blob in blobs:
                try:
                    candidates.append((True, _cells_from_blob(blob)))
                except Exception:
                    pass

        def anchors(cells):
            return sum(1 for s in cells if SEQ_RE.match(s))
        # 目标 tab 的候选优先，其中锚点最多者胜
        pool = [c for mt, c in candidates if mt] or [c for _, c in candidates]
        if pool:
            result["cells"] = max(pool, key=anchors)
        b.close()
    return result


# ══════════════════════ 行还原 ══════════════════════

_SHORT_CN = re.compile(r'^[一-龥A-Za-z]{2,6}$')


def parse_rows(cells: list) -> list:
    """
    把顺序文本流还原成行（单号为锚点）。
    车厂/人名区分：车厂名必然出现在“需求名称”长文本里（如 北汽-C66T-…），
    人名不会——先全局收集车厂集合，再据此归类，最后向下填充合并单元格。
    """
    groups, cur = [], None
    for s in cells:
        s = s.strip()
        if SEQ_RE.match(s):
            if cur:
                groups.append(cur)
            cur = {"seq": s, "items": []}
        elif cur is not None and s:
            cur["items"].append(s)
    if cur:
        groups.append(cur)

    # 第一遍：全局车厂集合 = 出现在本行长文本里的短中文串
    brands = set()
    for g in groups:
        longtext = " ".join(t for t in g["items"] if len(t) > 12)
        for t in g["items"]:
            if _SHORT_CN.match(t) and t in longtext \
                    and t not in LANG_SET and t not in STATUS_SET:
                brands.add(t)

    # 第二遍：逐行归类
    rows = []
    for g in groups:
        r = {"单号": g["seq"], "需求名称": "", "车厂": "", "语种": "",
             "云端or本地": "", "预计完成时间": "", "当前状态": "",
             "责任人": "", "备注": ""}
        extras = []
        for t in g["items"]:
            base = t.replace(" ", "")
            if not r["需求名称"] and len(t) > 15 and "-" in t:
                r["需求名称"] = t
            elif base in LANG_SET and not r["语种"]:
                r["语种"] = base
            elif t in CLOUD_SET and not r["云端or本地"]:
                r["云端or本地"] = t
            elif DATE_RE.match(t) and not r["预计完成时间"]:
                r["预计完成时间"] = t
            elif t in STATUS_SET and not r["当前状态"]:
                r["当前状态"] = t
            elif t in brands and not r["车厂"]:
                r["车厂"] = t
            elif _SHORT_CN.match(t) and t not in brands and not r["责任人"]:
                r["责任人"] = t
            else:
                extras.append(t)
        r["备注"] = " / ".join(extras)[:80]
        rows.append(r)

    # 第三遍：单号后缀权威推导 + 合并单元格向下填充
    try:
        from web_download import LANG_CN
    except Exception:
        LANG_CN = {}
    prev = {}
    for r in rows:
        if "-L-" in r["单号"]:
            r["云端or本地"] = "本地"
        elif "-C-" in r["单号"]:
            r["云端or本地"] = "云端"
        m = re.search(r'ASR-([A-Za-z_]+)$', r["单号"])
        if m:
            cn = LANG_CN.get(m.group(1).lower())
            if cn:
                r["语种"] = cn
        for col in ["车厂", "语种", "预计完成时间", "当前状态", "责任人"]:
            if not r[col] and prev.get(col):
                r[col] = prev[col]
        prev = {k: r[k] for k in ["车厂", "语种", "预计完成时间", "当前状态", "责任人"]}
    return rows


# ══════════════════════ 下载（复用 RMP 会话）══════════════════════

def _norm_seq(seq: str) -> str:
    """去掉 -L-/-C- 差异后的“同名”键：N-CUS-5967-L-ASR-th_th 与 -C- 版同键。"""
    return re.sub(r'-(?:L|C)-ASR', '-ASR', seq or '')


def _seq_num(seq: str):
    m = re.search(r'(\d{4,})', seq or '')
    return int(m.group(1)) if m else None


def _title_key(name: str) -> str:
    """需求名称去掉单号和空白后的可比键（NS 相邻单的名称除单号外相同）。"""
    if not name:
        return ''
    s = re.sub(r'(?:N-CUS|NS)-\d+(?:-[A-Za-z_]+)*', '', name)
    return re.sub(r'[\s　]+', '', s)


def dedup_cloud_local(rows: list) -> tuple:
    """
    云端/本地同名去重（本地已有时跳过云端；云端独有的保留）：
      规则1 同号：单号仅 -L-/-C- 不同（N-CUS-5967-L/-C-ASR-th_th）
      规则2 NS相邻号：NS- 单编号差 ±1，且【名称去单号后一致】或【车厂+语种一致】
             （双保险防误伤：相邻但不同语种/不同需求的不算同名）
    """
    locals_norm = set()
    locals_seqs = set()
    locals_ns = []
    for r in rows:
        seq = r["单号"]
        if "-L-" in seq:
            locals_norm.add(_norm_seq(seq))
            locals_seqs.add(seq)
            if seq.startswith("NS-"):
                locals_ns.append((_seq_num(seq), r.get("车厂", ""),
                                  r.get("语种", ""), _title_key(r.get("需求名称", ""))))
    kept, skipped = [], []
    for r in rows:
        seq = r["单号"]
        skip = False
        if "-C-" in seq:
            if _norm_seq(seq) in locals_norm:
                skip = True
            else:
                # 规则3（最硬证据）：需求名称里直接写着某个本地单号
                refs = re.findall(r'(?:N-CUS|NS)-\d+(?:-[A-Za-z_]+)*',
                                  r.get("需求名称", ""))
                if any(x in locals_seqs for x in refs):
                    skip = True
                elif seq.startswith("NS-"):
                    num = _seq_num(seq)
                    b, l = r.get("车厂", ""), r.get("语种", "")
                    tk = _title_key(r.get("需求名称", ""))
                    for n2, b2, l2, tk2 in locals_ns:
                        if num is None or n2 is None or abs(num - n2) != 1:
                            continue
                        if (tk and tk == tk2) or (b and l and b == b2 and l == l2):
                            skip = True
                            break
        if skip:
            skipped.append(seq)
        else:
            kept.append(r)
    return kept, skipped


def download_seqnos(seqnos: list, base: str = RMP_BASE) -> dict:
    from playwright.sync_api import sync_playwright
    if not os.path.exists(STATE):
        print("✗ 未找到登录会话，请先在界面点“登录”。", flush=True)
        return {"ok": 0, "fail": len(seqnos)}
    os.makedirs(INBOX, exist_ok=True)
    base = base.rstrip("/")
    search = base + "/rmp/v2/bug/requirement/search"
    attach = base + "/rmp/v2/bug/requirement/comment/getAttachList"
    ok = fail = 0
    with sync_playwright() as p:
        b = p.chromium.launch(headless=True)
        ctx = b.new_context(storage_state=STATE)
        req = ctx.request
        for i, seq in enumerate(seqnos, 1):
            m = re.search(r'\d{4,}', seq)
            if not m:
                fail += 1
                continue
            body = {"keyWord": m.group(), "purityStatus": "", "operatorStatus": "",
                    "requirementType": "", "status": "", "startTime": "", "endTime": "",
                    "cloudNativeFlag": "", "language": "", "createUser": "",
                    "pageNo": 1, "pageSize": 50}
            try:
                r = req.post(search, data=body, timeout=40000)
                items = (r.json() or {}).get("data", {}).get("content", []) or []
                hit = next((it for it in items if str(it.get("seqNo")) == seq), None)
                if not hit:
                    print(f"[{i}/{len(seqnos)}] {seq} 平台未找到，跳过", flush=True)
                    fail += 1
                    continue
                rid = str(hit.get("id"))
                atts = (req.get(f"{attach}?reqId={rid}&pageNo=1&pageSize=9999",
                                timeout=30000).json() or {}
                        ).get("data", {}).get("content", []) or []
                picked = None
                for kw in ["资源汇总", "逆规整后"]:
                    for a in atts:
                        if kw in str(a.get("name", "")) and str(a.get("size") or "0") != "0":
                            picked = a
                            break
                    if picked:
                        break
                if not picked:
                    print(f"[{i}/{len(seqnos)}] {seq} 无语料附件，跳过", flush=True)
                    fail += 1
                    continue
                data = req.get(picked["path"], timeout=120000).body()
                fname = re.sub(r'[\\/:*?"<>|]', "_", seq) + ".xlsx"
                with open(os.path.join(INBOX, fname), "wb") as f:
                    f.write(data)
                print(f"[{i}/{len(seqnos)}] {seq}  ↓ {len(data)//1024}KB", flush=True)
                ok += 1
            except Exception as e:
                print(f"[{i}/{len(seqnos)}] {seq} 失败：{e}", flush=True)
                fail += 1
        b.close()
    return {"ok": ok, "fail": fail}


# ══════════════════════ 主入口 ══════════════════════

def _latest_demand_tab(tabs) -> str:
    """从 tabs [(id,name,hidden)] 里挑最新的『YYYY-MM…需求单』tab 的 id。"""
    best = None
    for i, n, _h in tabs:
        m = re.search(r'(\d{4})\D+0?(\d{1,2}).*需求单', str(n))
        if m:
            key = (int(m.group(1)), int(m.group(2)))
            if best is None or key > best[0]:
                best = (key, i, n)
    return best[1] if best else ""


def main():
    ap = argparse.ArgumentParser(description="腾讯排期表读取/下载")
    ap.add_argument("cmd", choices=["tabs", "read", "sync"])
    ap.add_argument("--url", default=DEFAULT_DOC)
    ap.add_argument("--tab", default="")
    ap.add_argument("--user", default="", help="按责任人过滤")
    ap.add_argument("--lang", default="", help="按语种过滤")
    ap.add_argument("--brand", default="", help="按车厂过滤")
    ap.add_argument("--scope", default="all", choices=["all", "local", "cloud"])
    ap.add_argument("--download", action="store_true", help="下载命中的单号语料")
    ap.add_argument("--no-dedup", action="store_true",
                    help="关闭云端/本地同名去重（默认开启：本地有同名时跳过云端）")
    args = ap.parse_args()

    print(f"读取排期表：{args.url}" + (f"  tab={args.tab}" if args.tab else ""), flush=True)
    doc = fetch_doc(args.url, args.tab)

    if args.cmd == "tabs":
        vis = [(i, n) for i, n, h in doc["tabs"] if not h]
        hid = [(i, n) for i, n, h in doc["tabs"] if h]
        print(f"可见 tab（{len(vis)}）:", flush=True)
        for i, n in vis:
            print(f"  {i}  {n}", flush=True)
        print(f"（另有 {len(hid)} 个隐藏 tab）", flush=True)
        return 0

    if args.cmd == "sync":
        # 「刷新排期」：从文档最新需求单 tab 拉取，覆盖更新本地 排期.xlsx 的车厂/语种
        tab = args.tab or _latest_demand_tab(doc["tabs"])
        if not tab:
            print("✗ 找不到需求单 tab（可用 --tab 指定）", flush=True)
            return 1
        tname = next((n for i, n, _h in doc["tabs"] if i == tab), tab)
        print(f"从文档刷新排期：tab = {tname}", flush=True)
        sdoc = fetch_doc(args.url, tab)
        srows = parse_rows(sdoc["cells"])
        if not srows:
            print("✗ 该 tab 没解析到排期行。", flush=True)
            return 1
        recs = [(r["单号"], r["车厂"], r["语种"], r.get("预计完成时间", ""))
                for r in srows if r.get("单号")]
        try:
            import web_download as wd
            upd, add = wd.upsert_schedule(recs)
            print(f"✓ 刷新完成：更新 {upd} 行，新增 {add} 行（车厂/语种以文档为准）"
                  f"  ->  排期.xlsx", flush=True)
            return 0
        except Exception as e:
            print(f"✗ 写排期失败：{e}", flush=True)
            return 1

    rows = parse_rows(doc["cells"])
    if not rows:
        print("✗ 没解析到任何行（tab 可能不是排期格式，或文档无法访问）", flush=True)
        return 1
    # 云端/本地同名去重（先于任何过滤，本地同名存在与否必须以全集为准，
    # 否则责任人/语种/车厂过滤先删掉本地孪生 → 云端单不再被去重）
    if not args.no_dedup:
        rows, dup_skipped = dedup_cloud_local(rows)
        if dup_skipped:
            print(f"\n云端/本地同名去重：跳过 {len(dup_skipped)} 个云端单（本地已有同名）：", flush=True)
            for s in dup_skipped:
                print(f"  · {s}", flush=True)
    def _fz(needle, *fields):
        n = str(needle).lower().replace(" ", "")
        return any(n in str(f or "").lower().replace(" ", "") for f in fields)
    if args.user:
        rows = [r for r in rows if args.user in r["责任人"] or args.user in r["备注"]]
    if args.lang:
        rows = [r for r in rows if _fz(args.lang, r["语种"], r["单号"])]
    if args.brand:
        rows = [r for r in rows if _fz(args.brand, r["车厂"])]
    if args.scope == "local":
        rows = [r for r in rows if "-L-" in r["单号"]]
    elif args.scope == "cloud":
        rows = [r for r in rows if "-C-" in r["单号"]]

    print(f"\n命中 {len(rows)} 行：", flush=True)
    for r in rows:
        print(f"  {r['单号']:28s} {r['车厂']:6s} {r['语种']:6s} "
              f"{r['云端or本地']:4s} {r['当前状态']:5s} {r['责任人']}", flush=True)

    # 写排期
    try:
        import web_download as wd
        sched = [(r["单号"], r["车厂"], r["语种"] or wd._lang_cn(r["单号"]),
                  r["预计完成时间"]) for r in rows]
        added = wd.update_schedule(sched)
        print(f"\n✓ 排期表：自动新增 {added} 行（已存在不重复）", flush=True)
    except Exception as e:
        print(f"[warn] 排期更新失败：{e}", flush=True)

    if args.download and rows:
        print("\n开始下载命中的语料…", flush=True)
        res = download_seqnos([r["单号"] for r in rows])
        print(f"✓ 下载完成：成功 {res['ok']}，失败/跳过 {res['fail']}", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
