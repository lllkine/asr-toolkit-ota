# -*- coding: utf-8 -*-
"""
引擎查找/下载（xfchat「2026智能座舱版本更新记录」表）：
  python xf_engine.py login                       # 弹浏览器，i讯飞扫码登录一次
  python xf_engine.py tabs                        # 列出表格的所有 tab
  python xf_engine.py list [--tab 本地多语种_听写] [--lang 泰语] [--brand 极氪]
  python xf_engine.py get  --lang 泰语 --brand 极氪 [--tab ...]
                                                  # 取最新一条，打印货架路径
  python xf_engine.py get  ... --to rdg           # 用 rftctl warehouse 传到内网
  python xf_engine.py get  ... --local            # 下载到本机 _engines/（需先 login-devops）
  python xf_engine.py login-devops                # 弹浏览器登 E3/devops（下载引擎用）
"""
import os
import re
import sys
import csv
import io
import argparse
import subprocess


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

SESSION_DIR = os.path.join(APP, "_session")
XF_STATE = os.path.join(SESSION_DIR, "xf_state.json")
DEVOPS_STATE = os.path.join(SESSION_DIR, "devops_state.json")
DEVOPS_PROFILE = os.path.join(SESSION_DIR, "devops_profile")  # 持久化配置目录
ENGINES_DIR = os.path.join(APP, "_engines")


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


XF_URL = _endpoint("XF_SHEET_URL")
DEFAULT_TAB = "本地多语种_听写"
DEVOPS_BASE = _endpoint("DEVOPS_BASE")                       # 如 http://<devops-host>
DEVOPS_HOST = DEVOPS_BASE.split("://")[-1].split("/")[0]    # 提取主机名用于登录检测
RFTCTL_DEFAULT = _endpoint("RFTCTL_PATH")


def _find_rftctl() -> str:
    import shutil as _sh
    p = os.environ.get('RFTCTL_PATH')
    if p and os.path.exists(p):
        return p
    found = _sh.which('rftctl')
    if found:
        return found
    return RFTCTL_DEFAULT if os.path.exists(RFTCTL_DEFAULT) else ''


# ══════════════════ 登录 ══════════════════

def login() -> int:
    from playwright.sync_api import sync_playwright
    os.makedirs(SESSION_DIR, exist_ok=True)
    print("正在打开浏览器…请用 i讯飞 手机App 扫码登录，登录后【关闭浏览器窗口】。", flush=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = browser.new_context(no_viewport=True)
        page = ctx.new_page()
        try:
            page.goto(XF_URL, wait_until="domcontentloaded", timeout=60000)
        except Exception:
            pass
        i = 0
        while True:
            try:
                page.wait_for_timeout(1000)
                if not ctx.pages:
                    break
                i += 1
                if i % 5 == 0:
                    ctx.storage_state(path=XF_STATE)
            except Exception:
                break
        try:
            ctx.storage_state(path=XF_STATE)
        except Exception:
            pass
        try:
            browser.close()
        except Exception:
            pass
    ok = os.path.exists(XF_STATE)
    print("✓ 会话已保存。" if ok else "✗ 未捕获会话，请重试。", flush=True)
    return 0 if ok else 1


def _devops_logged_in(url: str) -> bool:
    """URL 落在 devops 域且不在登录/SSO 页 = 已登录。"""
    u = (url or "").lower()
    if "sso" in u or "login" in u or "cas" in u or "/auth" in u:
        return False
    return bool(DEVOPS_HOST) and DEVOPS_HOST in u


def login_devops() -> int:
    """登 E3/devops（引擎本地下载用）。集团 SSO：账号密码或扫码。
    用持久化配置目录，登录票据落盘到 profile，下载复用同一 profile —— 比
    另开无头浏览器重放 storage_state 稳（SSO 会话票据不易丢）。"""
    from playwright.sync_api import sync_playwright
    os.makedirs(DEVOPS_PROFILE, exist_ok=True)
    url = DEVOPS_BASE + "/ipackage"
    print("正在打开浏览器…请登录 E3 平台（集团账号/扫码）。", flush=True)
    print("登录成功后工具会自动识别，届时你会看到「✓ 已检测到已登录」，再关闭窗口即可。", flush=True)
    detected = False
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            DEVOPS_PROFILE, headless=False, accept_downloads=True,
            no_viewport=True, args=["--start-maximized"])
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=60000)
        except Exception:
            pass
        while True:
            try:
                page.wait_for_timeout(1000)
                if not ctx.pages:
                    break
                cur = ctx.pages[-1].url
                if not detected and _devops_logged_in(cur):
                    detected = True
                    print("✓ 已检测到已登录（制品库）。可以关闭浏览器窗口了。", flush=True)
                    try:
                        ctx.storage_state(path=DEVOPS_STATE)
                    except Exception:
                        pass
            except Exception:
                break
        try:
            ctx.storage_state(path=DEVOPS_STATE)
        except Exception:
            pass
        try:
            ctx.close()
        except Exception:
            pass
    if detected:
        print("✓ devops 会话已保存（已确认登录）。", flush=True)
        return 0
    print("✗ 未检测到成功登录制品库（可能卡在 SSO 登录页 / 需要内网或二次验证）。请重试。", flush=True)
    return 1


# ══════════════════ 读表（剪贴板法）══════════════════

def _active_tab(page) -> str:
    """当前选中的 tab 名（读 active/selected 类名或 aria-selected）。读不出返回 ''。"""
    try:
        return page.evaluate("""() => {
            const els = Array.from(document.querySelectorAll(
                '[class*="tab"],[role="tab"],[class*="sheet"]'));
            const a = els.find(e =>
                e.getAttribute('aria-selected') === 'true' ||
                /active|selected|current/i.test(e.className || ''));
            return a ? (a.textContent || '').trim() : '';
        }""") or ""
    except Exception:
        return ""


def _resolve_tab(page, tab_name: str) -> bool:
    """切到指定 tab。三次重试：页面渲染有快有慢，一次点不中很常见；
    而且本来就停在目标 tab 上时（默认 tab 就是它）根本不需要点。"""
    if not tab_name:
        return True
    for attempt in range(3):
        cur = _active_tab(page)
        if cur and tab_name in cur:          # 已经在目标 tab 上，不用点
            return True
        try:
            page.get_by_text(tab_name, exact=True).first.click(timeout=6000)
            page.wait_for_timeout(4000)
            return True
        except Exception:
            if attempt < 2:
                page.wait_for_timeout(3000)  # 多半是还没渲染出来，等一会儿再试
    print(f"[warn] 未找到 tab「{tab_name}」（已重试 3 次），当前 tab：{_active_tab(page) or '未知'}",
          flush=True)
    return False


def fetch_sheet(tab: str = "", auto_login: bool = True) -> list:
    """打开表格 → 切 tab → 全选复制 → 返回 TSV 行列表。
    没登录 / 会话过期时：直接弹出扫码登录窗口，登完自动重来一次，不用用户自己找入口。"""
    from playwright.sync_api import sync_playwright
    if not os.path.exists(XF_STATE):
        if not auto_login:
            print("✗ 未登录 xfchat。", flush=True)
            return []
        print("! 尚未登录 xfchat，正在打开登录窗口，请用 i讯飞 App 扫码…", flush=True)
        if login() != 0:
            return []
        return fetch_sheet(tab, auto_login=False)

    expired = False
    rows = []
    with sync_playwright() as p:
        b = p.chromium.launch(headless=True)
        ctx = b.new_context(storage_state=XF_STATE,
                            viewport={"width": 1600, "height": 1000},
                            permissions=["clipboard-read", "clipboard-write"])
        page = ctx.new_page()
        try:
            page.goto(XF_URL, wait_until="domcontentloaded", timeout=60000)
        except Exception as e:
            print(f"[warn] {e}", flush=True)
        page.wait_for_timeout(11000)
        if "login" in page.url:
            expired = True
            b.close()
        else:
            rows = _copy_sheet_rows(page, tab)
            b.close()

    if not expired:
        return rows

    # 过期分支必须等 sync_playwright 退出后再走：login() 自己也开 playwright，不能嵌套
    try:
        os.remove(XF_STATE)          # 清掉过期票据，避免下次又拿它去撞登录页
    except OSError:
        pass
    if not auto_login:
        print("✗ xfchat 会话已失效，且自动重新登录后仍未通过。请点「引擎登录」重试。", flush=True)
        return []
    print("! xfchat 会话已失效，正在打开登录窗口，请用 i讯飞 App 扫码；"
          "登录后关闭浏览器窗口即可，工具会自动接着查。", flush=True)
    if login() != 0:
        return []
    return fetch_sheet(tab, auto_login=False)


def _copy_one(page) -> str:
    """点一下表格 → 全选 → 复制 → 读剪贴板。"""
    page.mouse.click(700, 500)
    page.wait_for_timeout(400)
    page.keyboard.press("Control+A")
    page.wait_for_timeout(400)
    page.keyboard.press("Control+C")
    page.wait_for_timeout(2000)
    try:
        return page.evaluate("navigator.clipboard.readText()") or ""
    except Exception as e:
        print(f"✗ 读表失败（剪贴板不可读）：{e}", flush=True)
        return ""


def _select_all_after_full_load(page) -> str:
    """表格是懒加载的：不滚到底就直接全选复制，只能拿到已渲染的前几百行——
    最近上架的引擎（表是按时间升序排的，新的在最后）全部读不到，
    于是"取最新"取到的其实是几个月前的旧引擎。
    所以：先反复滚到底，直到复制到的行数不再增长，再取最终结果。"""
    txt = _copy_one(page)
    last = len(txt.splitlines())
    for _round in range(8):
        for _ in range(15):                 # 一轮猛滚，触发后续分页加载
            page.keyboard.press("Control+End")
            page.wait_for_timeout(500)
            page.mouse.wheel(0, 20000)
            page.wait_for_timeout(300)
        page.wait_for_timeout(2000)
        txt = _copy_one(page)
        n = len(txt.splitlines())
        if n <= last:                       # 不再增长 = 已经到底
            break
        print(f"  [读表] 已加载 {n} 行…", flush=True)
        last = n
    print(f"  [读表] 共 {last} 行", flush=True)
    return txt


def _copy_sheet_rows(page, tab: str) -> list:
    """在已登录的页面上：切 tab → 全选复制 → TSV 行。"""
    tab_ok = _resolve_tab(page, tab)
    if tab and not tab_ok:
        # 没切中 tab 就复制，读到的是别的 tab 的数据，后面还会在错误数据里模糊匹配，
        # 最后可能把别的引擎的版本号/货架地址群发出去。宁可不给结果。
        print(f"✗ 未能切到 tab「{tab}」，为避免读到错误 tab 的数据，本次不返回结果。"
              f"请确认 tab 名是否正确（可用 tabs 命令列出）。", flush=True)
        return []
    txt = _select_all_after_full_load(page)
    if not txt:
        # 空串≠"表里没有匹配记录"。不说清楚的话，上层会打成"没有匹配的引擎记录"，
        # 用户以为引擎还没上架，实际是读表挂了（页面没加载完/tab 没切中/剪贴板权限）。
        print("✗ 读表失败：没从表格里复制到任何内容（页面可能没加载完，或剪贴板被拒）。"
              "这不代表『引擎不存在』，请重试或重新登录。", flush=True)
        return []
    return list(csv.reader(io.StringIO(txt), delimiter="\t"))


def parse_engines(rows: list) -> list:
    """TSV → 引擎记录。列：V R M 语种 归属 描述 上架同学 上架时间 货架位置

    以「货架 URL」这一列为锚点倒推各字段，不按固定下标取：
    表格复制出来时前面可能多一列（序号/复选框），按 r[0]=V 硬取会整体错位一格，
    把「语种」读成版本号、「归属」读成语种，匹配和展示全乱。"""
    out = []
    for r in rows:
        ui = -1
        for i in range(len(r) - 1, -1, -1):
            if str(r[i] or "").strip().lower().startswith("http"):
                ui = i
                break
        if ui < 0:
            continue

        def cell(idx, _r=r):
            return str(_r[idx]).strip() if 0 <= idx < len(_r) else ""

        desc_full = cell(ui - 3)
        date_t = _parse_date(cell(ui - 1)) or _row_date(r)
        rec = {
            "V": cell(ui - 8), "R": cell(ui - 7), "M": cell(ui - 6),
            "语种": cell(ui - 5), "归属": cell(ui - 4),
            "描述": desc_full[:60], "_描述全": desc_full,
            # 归一成 2026.05.12，补零后字符串序 == 时间序，肉眼和排序都不会再错
            "上架时间": ("%04d.%02d.%02d" % date_t) if date_t else "",
            "_日期": date_t,                      # 真正用来排序的元组，解析不出为 None
            "货架": cell(ui),
        }
        if rec["语种"]:
            out.append(rec)
    return out


_DATE_RE = re.compile(r'(20\d\d)\s*[.\-/年]\s*(\d{1,2})\s*[.\-/月]\s*(\d{1,2})')


def _parse_date(s):
    """'2026.5.12' / '2026-05-12' / '2026/5/12 10:30' / '2026年5月12日' → (2026,5,12)。
    解析不了返回 None。"""
    m = _DATE_RE.search(str(s or ''))
    if not m:
        return None
    y, mo, d = (int(x) for x in m.groups())
    if not (1 <= mo <= 12 and 1 <= d <= 31):
        return None
    return (y, mo, d)


def _row_date(r):
    """取「上架时间」。表头列位是 …上架同学(6) 上架时间(7) 货架(8)，优先按列位取；
    列位对不上时退回整行里【最后】一个像日期的单元格 —— 不能取第一个：
    行首的版本/计划等列也常常长得像日期，一取就取错，排序自然选错引擎。"""
    if len(r) > 7:
        d = _parse_date(r[7])
        if d:
            return d
    for c in reversed(r):
        d = _parse_date(c)
        if d:
            return d
    return None


def _match(rec, lang, brand):
    """宽松匹配：分开填（语种=泰语 车厂=极氪）或连写（极氪泰语）都能命中。"""
    combo = rec["归属"] + rec["语种"]
    ok = True
    if lang:
        ok = ok and (lang in rec["语种"] or lang in combo)
    if brand:
        ok = ok and (brand in rec["归属"] or brand in combo)
    return ok


def _fuzzy_match(rec, query: str) -> bool:
    """全文模糊：所有字段拼成一串，空格/逗号分隔的关键词全部命中即匹配（忽略大小写）。"""
    hay = "".join(str(v) for v in rec.values()).lower().replace(" ", "")
    toks = [t.lower().replace(" ", "") for t in re.split(r'[\s,，、/;；]+', query) if t.strip()]
    return bool(toks) and all(t in hay for t in toks)


def search(recs: list, lang: str, brand: str) -> tuple:
    """两级搜索：精确字段匹配优先；无结果时降级全文模糊。返回 (结果, 模式)。"""
    strict = [r for r in recs if _match(r, lang, brand)]
    if strict:
        return strict, "精确"
    query = f"{lang} {brand}".strip()
    if not query:
        return [], "精确"
    return [r for r in recs if _fuzzy_match(r, query)], "模糊"


# ══════════════════ 下载 / 传输 ══════════════════

def repo_path_of(url: str) -> str:
    m = re.search(r'[?&]repo=([^&]+)', url)
    return m.group(1) if m else ""


def _engine_dir_of(rec: dict) -> str:
    """引擎 zip 的本地落地目录（与 _download_rec_on_page 的命名保持一致）。"""
    sub = re.sub(r'[:*?"<>|]', "_", f"{rec['归属']}{rec['语种']}_{rec['上架时间']}")
    return os.path.join(ENGINES_DIR, sub)


def _zips_in(d: str) -> list:
    if not os.path.isdir(d):
        return []
    return [os.path.join(d, f) for f in os.listdir(d) if f.lower().endswith('.zip')]


def send_local_engine(rec: dict, networks: str) -> int:
    """下载引擎 zip 到本机，再用 rftctl『local submit』把文件发到各网段。

    不用 warehouse submit：那条路是让服务端自己从制品库转运，一直 502，指望不上。
    先下到本地再传文件，链路上少一个坏掉的服务。"""
    rft = _find_rftctl()
    if not rft:
        print("✗ 未找到 rftctl。", flush=True)
        return 1

    dest = _engine_dir_of(rec)
    zips = _zips_in(dest)
    if zips:
        print(f"已有本地引擎包，跳过下载：{dest}", flush=True)
    else:
        print("本地还没有这个引擎包，先下载（会弹浏览器登录 E3）…", flush=True)
        if login_and_download([rec]) != 0:
            print("✗ 引擎下载失败，无法传输。", flush=True)
            return 1
        zips = _zips_in(dest)
    if not zips:
        print(f"✗ {dest} 里没有 zip，无法传输。", flush=True)
        return 1

    nets = [n.strip() for n in str(networks).split(',') if n.strip()]
    title = f"引擎传输_{rec['归属']}{rec['语种']}_{rec['上架时间'].replace('.', '')}"
    remark = f"{rec['上架时间']}引擎传输{rec['归属']}{rec['语种']}"
    print(f"传输 {len(zips)} 个文件到 {nets}：{', '.join(os.path.basename(z) for z in zips)}",
          flush=True)

    failed = []
    for net in nets:
        cmd = [rft, "local", "submit", "--receive", net,
               "--file", dest, "--remark", remark, "--title", title]
        print(f"  -> [{net}] 提交中… {' '.join(cmd)}", flush=True)
        r = subprocess.run(cmd, text=True, encoding="utf-8", errors="replace",
                           capture_output=True)
        detail = ((r.stdout or '') + (r.stderr or '')).strip()
        if detail:
            print(f"     rftctl: {detail}", flush=True)
        if r.returncode == 0:
            print(f"  ✓ [{net}] 传输完成。", flush=True)
        else:
            failed.append(net)
            print(f"  ✗ [{net}] 传输失败（exit {r.returncode}）。", flush=True)
    if failed:
        print(f"✗ 失败网段：{failed}。本地包仍在 {dest}，可稍后重试。", flush=True)
        return 1
    print(f"✓ 全部网段传输完成：{nets}", flush=True)
    return 0


def send_to_network(rec: dict, network: str) -> int:
    rft = _find_rftctl()
    if not rft:
        print("✗ 未找到 rftctl。", flush=True)
        return 1
    repo = repo_path_of(rec["货架"])
    if not repo:
        print("✗ 货架链接里没有 repo 路径。", flush=True)
        return 1
    title = f"引擎传输_{rec['归属']}{rec['语种']}_{rec['上架时间'].replace('.', '')}"
    cmd = [rft, "warehouse", "submit", "--title", title,
           "--remark", f"{rec['上架时间']}引擎传输{rec['归属']}{rec['语种']}",
           "--target-network", network, "--files", repo]
    print("执行:", " ".join(cmd), flush=True)
    # 抓住 rftctl 自己的输出：不抓的话失败时只剩一个 exit 1，没人知道它到底在抱怨什么
    r = subprocess.run(cmd, text=True, encoding="utf-8", errors="replace",
                       capture_output=True)
    detail = ((r.stdout or '') + (r.stderr or '')).strip()
    if r.returncode == 0:
        if detail:
            print(detail, flush=True)
        print("✓ 制品库传输流程已提交。", flush=True)
    else:
        if detail:
            print(f"  rftctl 原话：{detail}", flush=True)
        print(f"✗ 提交失败（exit {r.returncode}）。若提示需要具体文件路径，"
              f"可在货架页面找到文件名后重试：--files {repo}/<文件名>", flush=True)
    return r.returncode


def _click_download_control(page):
    """勾选文件后，点击出现的「下载」按钮。返回是否点到。"""
    # iPackage 勾选后会出现下载入口（文字/图标/按钮），逐个选择器尝试可见的
    cands = [
        page.get_by_role("button", name=re.compile("下载")),
        page.locator('[title*="下载"]'),
        page.locator('[aria-label*="下载"]'),
        page.get_by_text("下载", exact=True),
        page.locator('button:has-text("下载"), a:has-text("下载"), span:has-text("下载")'),
    ]
    for loc in cands:
        try:
            n = loc.count()
        except Exception:
            n = 0
        for i in range(n):
            el = loc.nth(i)
            try:
                if el.is_visible() and el.is_enabled():
                    el.click(timeout=5000)
                    return True
            except Exception:
                continue
    return False


def _download_rec_on_page(page, ctx, rec: dict) -> int:
    """在一个已登录的页面上下载单个引擎货架里的 zip：勾选复选框→点下载。返回文件数。"""
    os.makedirs(ENGINES_DIR, exist_ok=True)
    sub = re.sub(r'[\\/:*?"<>|]', "_", f"{rec['归属']}{rec['语种']}_{rec['上架时间']}")
    dest_dir = os.path.join(ENGINES_DIR, sub)
    os.makedirs(dest_dir, exist_ok=True)
    got = 0
    try:
        page.goto(rec["货架"], wait_until="domcontentloaded", timeout=60000)
    except Exception as e:
        print(f"[warn] {e}", flush=True)
    page.wait_for_timeout(9000)
    if not _devops_logged_in(page.url):
        print("✗ 该页面被踢到登录页，会话可能已失效。", flush=True)
        return 0

    # 找出所有文件行（tr），挑出 .zip（引擎包）
    rows = page.query_selector_all("tr")
    targets = []  # (row_element, filename)
    for tr in rows:
        try:
            txt = (tr.inner_text() or "").strip()
        except Exception:
            continue
        low = txt.lower()
        if any(ext in low for ext in [".zip", ".tar", ".tgz", ".gz", ".7z"]):
            fn = ""
            for tok in re.split(r'\s+', txt):
                if re.search(r'\.(zip|tar|tgz|gz|7z)$', tok, re.I):
                    fn = tok
                    break
            targets.append((tr, fn or txt[:40]))
    print(f"发现 {len(targets)} 个压缩包文件行", flush=True)

    for tr, fn in targets:
        # 复选框是自定义样式（Element-UI），真正的 <input> 隐藏，需点可见的 wrapper
        clicked = False
        for sel in ['.el-checkbox__inner', '.el-checkbox', '[class*="checkbox"]',
                    'td:first-child', 'input[type=checkbox]']:
            el = tr.query_selector(sel)
            if not el:
                continue
            try:
                el.click(force=True, timeout=4000)
                clicked = True
                break
            except Exception:
                continue
        if not clicked:
            print(f"  [跳过] {fn}: 勾选失败（找不到可点的复选框）", flush=True)
            continue
        page.wait_for_timeout(1500)
        try:
            with page.expect_download(timeout=300000) as dl_info:
                if not _click_download_control(page):
                    # dump 当前可见按钮，方便定位下载入口
                    try:
                        btns = page.evaluate(
                            r"""() => Array.from(document.querySelectorAll('button,a,span,i,[role=button]'))
                                 .filter(e => e.offsetParent!==null)
                                 .map(e => (e.getAttribute('title')||e.innerText||e.getAttribute('class')||'').trim())
                                 .filter(t => t && t.length<25)
                                 .filter((v,i,a)=>a.indexOf(v)===i).slice(0,40)""")
                        print(f"  [调试] 勾选后可见控件: {btns}", flush=True)
                    except Exception:
                        pass
                    raise RuntimeError("未找到下载按钮")
            dl = dl_info.value
            path = os.path.join(dest_dir, dl.suggested_filename or fn)
            dl.save_as(path)
            sz = os.path.getsize(path) if os.path.exists(path) else 0
            print(f"  ↓ {dl.suggested_filename or fn} ({sz//1024//1024}MB)", flush=True)
            got += 1
        except Exception as e:
            print(f"  [失败] {fn}: {str(e)[:120]}", flush=True)
        # 取消勾选，避免影响下一个（原来写的是 cb，这变量根本不存在 → 每轮抛 NameError
        # 被 except 吞掉，复选框其实从没被取消过：货架里有多个包时会累积勾选，
        # 下载按钮作用于"已勾选的全部"，而我们只接住第一个 download 事件）
        try:
            if el is not None:
                el.click(force=True, timeout=3000)
                page.wait_for_timeout(400)
        except Exception as e:
            print(f"  [warn] {fn}: 取消勾选失败（{str(e)[:60]}），"
                  f"下一个包可能被连带下载", flush=True)

    if not targets:
        shot = os.path.join(dest_dir, "货架页面.png")
        try:
            page.screenshot(path=shot, full_page=True)
        except Exception:
            pass
        print(f"未发现压缩包文件，已截图货架页面 -> {shot}", flush=True)
        print(f"货架地址：{rec['货架']}", flush=True)
    print(f"  → {got} 个文件 -> {dest_dir}", flush=True)
    return got


def download_local(rec: dict) -> int:
    """打开货架页面，找文件下载链接下到 _engines/。复用 devops 持久化 profile。"""
    from playwright.sync_api import sync_playwright
    if not os.path.isdir(DEVOPS_PROFILE):
        print("✗ 未登录 devops，请先运行 login-devops（界面上点「登录E3」）。", flush=True)
        return 2
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            DEVOPS_PROFILE, headless=True, accept_downloads=True,
            viewport={"width": 1600, "height": 1000})
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        # 先探一下会话是否还在（SSO 票据可能已随浏览器关闭失效）
        try:
            page.goto(DEVOPS_BASE + "/ipackage",
                      wait_until="domcontentloaded", timeout=60000)
        except Exception:
            pass
        page.wait_for_timeout(6000)
        if not _devops_logged_in(page.url):
            print("✗ devops 会话已失效（SSO 票据随浏览器关闭失效）。"
                  "请用「登录并下载」一次性完成。", flush=True)
            ctx.close()
            return 2
        got = _download_rec_on_page(page, ctx, rec)
        ctx.close()
    return 0 if got else 3


def login_and_download(picked: list) -> int:
    """一次性流程：开有头浏览器→用户登录→检测成功→在同一活会话里顺序下载所有引擎。
    这样避开「SSO 会话票据随浏览器关闭失效」的问题（无需重开无头浏览器）。"""
    from playwright.sync_api import sync_playwright
    os.makedirs(DEVOPS_PROFILE, exist_ok=True)
    if not picked:
        print("✗ 没有待下载的引擎清单。", flush=True)
        return 1
    print(f"待下载 {len(picked)} 个引擎。", flush=True)
    print("", flush=True)
    print("┌────────────────────────────────────────────────────────────┐", flush=True)
    print("│  即将打开浏览器，请登录 E3 平台（集团账号 / 扫码）。          │", flush=True)
    print("│  登录成功后会【自动开始下载】，引擎包上百 MB，需要几分钟。    │", flush=True)
    print("│  ★ 请一直开着浏览器，等日志出现「全部完成」再关！             │", flush=True)
    print("│    E3 的登录票据是会话级的，中途关掉窗口下载会直接中断。      │", flush=True)
    print("└────────────────────────────────────────────────────────────┘", flush=True)
    print("", flush=True)
    total_files = 0
    ok_engines = 0
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            DEVOPS_PROFILE, headless=False, accept_downloads=True,
            no_viewport=True, args=["--start-maximized"])
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        try:
            page.goto(DEVOPS_BASE + "/ipackage",
                      wait_until="domcontentloaded", timeout=60000)
        except Exception:
            pass
        # 等待登录（最多 5 分钟）
        waited = 0
        while waited < 300:
            page.wait_for_timeout(1000)
            waited += 1
            if not ctx.pages:
                print("✗ 浏览器被提前关闭，未开始下载。", flush=True)
                return 1
            if _devops_logged_in(ctx.pages[-1].url):
                break
        else:
            print("✗ 5 分钟内未检测到登录成功（可能卡在 SSO 页）。", flush=True)
            ctx.close()
            return 1
        print("✓ 已检测到登录，开始下载（请保持窗口打开）…", flush=True)
        dl_page = ctx.new_page()
        for i, rec in enumerate(picked, 1):
            print(f"═══ [{i}/{len(picked)}] {rec['归属']}{rec['语种']} {rec['上架时间']} ═══",
                  flush=True)
            got = _download_rec_on_page(dl_page, ctx, rec)
            total_files += got
            if got:
                ok_engines += 1
            print(flush=True)
        try:
            ctx.storage_state(path=DEVOPS_STATE)
        except Exception:
            pass
        ctx.close()
    print(f"全部完成：成功 {ok_engines}/{len(picked)} 个引擎，共 {total_files} 个文件 "
          f"-> {ENGINES_DIR}", flush=True)
    return 0 if ok_engines else 3


# ══════════════════ 主入口 ══════════════════

def main():
    ap = argparse.ArgumentParser(description="引擎查找/下载")
    ap.add_argument("cmd", choices=["login", "login-devops", "tabs", "list", "get"])
    ap.add_argument("--tab", default=DEFAULT_TAB)
    ap.add_argument("--lang", default="")
    ap.add_argument("--brand", default="")
    # --to 接受逗号分隔的多个网段（如 rdg,dtn）：默认走「下载 zip 再 local submit」
    ap.add_argument("--to", default="", help="接收网段，逗号分隔，如 rdg 或 rdg,dtn")
    ap.add_argument("--local", action="store_true", help="只下载到本机，不传输")
    ap.add_argument("--warehouse", action="store_true",
                    help="老路：让服务端从制品库直接转运（warehouse submit，常年 502）")
    args = ap.parse_args()

    if args.cmd == "login":
        sys.exit(login())
    if args.cmd == "login-devops":
        sys.exit(login_devops())

    if args.cmd == "tabs":
        # tab 名在页面顶部，直接列出（读 DOM 文本）
        from playwright.sync_api import sync_playwright
        if not os.path.exists(XF_STATE):
            print("! 尚未登录 xfchat，正在打开登录窗口，请用 i讯飞 App 扫码…", flush=True)
            if login() != 0:
                sys.exit(2)
        with sync_playwright() as p:
            b = p.chromium.launch(headless=True)
            ctx = b.new_context(storage_state=XF_STATE, viewport={"width": 1600, "height": 1000})
            page = ctx.new_page()
            page.goto(XF_URL, wait_until="domcontentloaded", timeout=60000)
            page.wait_for_timeout(10000)
            names = page.evaluate(
                """() => Array.from(document.querySelectorAll('[class*="tab"]'))
                     .map(e => e.textContent.trim()).filter(t => t && t.length < 30)""")
            b.close()
        seen = []
        for n in names:
            if n not in seen:
                seen.append(n)
        print("tabs:", flush=True)
        for n in seen[:30]:
            print("  " + n, flush=True)
        sys.exit(0)

    rows = fetch_sheet(args.tab)
    recs, mode = search(parse_engines(rows), args.lang, args.brand)
    if not recs:
        print("✗ 没有匹配的引擎记录（精确和模糊都未命中）。", flush=True)
        sys.exit(1)
    # 按真实日期排序（原来是按字符串排：'2026-05-12' < '2026.3.20'、'2026.12.1' < '2026.5.12'
    # 都会排反，get 取 recs[-1] 就会下到旧引擎）。日期解析不出来的排最前，不许当"最新"。
    recs.sort(key=lambda r: r.get("_日期") or (0, 0, 0))
    undated = [r for r in recs if not r.get("_日期")]

    if args.cmd == "list":
        print(f"匹配 {len(recs)} 条（{mode}匹配，旧→新）：", flush=True)
        for r in recs:
            print(f"  {(r['上架时间'] or '(无日期)'):12s} {r['归属']:10s} {r['语种']:8s} "
                  f"V{r['V']} R{r['R']} M{r['M']}  {r['描述'][:30]}", flush=True)
        sys.exit(0)

    # get：取最新
    rec = recs[-1]
    if undated:
        print(f"! {len(undated)} 条记录的上架时间读不出来，已排除在『最新』判定之外："
              f"{'、'.join((r['归属'] + r['语种'] + ' V' + r['V']) for r in undated[:5])}",
              flush=True)
    if not rec.get("_日期"):
        print("✗ 匹配到的记录都没有可识别的上架时间，无法判断哪个最新。"
              "请用 list 命令人工确认后再下载。", flush=True)
        sys.exit(1)
    same_day = [r for r in recs if r.get("_日期") == rec["_日期"]]
    if len(same_day) > 1:
        print(f"! 有 {len(same_day)} 条同为最新日期 {rec['上架时间']}，取表中最后一条。"
              f"如不对请用 list 确认。", flush=True)
    print("最新引擎：", flush=True)
    for k in ["上架时间", "归属", "语种", "V", "R", "M", "描述"]:
        print(f"  {k}: {rec[k]}", flush=True)
    print(f"  货架: {rec['货架']}", flush=True)
    print(f"  repo: {repo_path_of(rec['货架'])}", flush=True)

    code = 0
    if args.to:
        if args.warehouse:
            code = send_to_network(rec, args.to) or code       # 老路：服务端转运（常年 502）
        else:
            # 默认：先把 zip 下到本机，再用 local submit 传文件
            code = send_local_engine(rec, args.to) or code
    elif args.local:
        # 用「登录即下载」活会话流程：E3 的 SSO 票据是会话级的，
        # 关浏览器就失效，必须在同一活浏览器里登录后立即下载。
        code = login_and_download([rec]) or code
    sys.exit(code)


if __name__ == "__main__":
    main()
