# -*- coding: utf-8 -*-
"""
ASR 语料处理流水线 —— 图形界面 (PySide6)
引导式：① 下载语料  ② 处理  ③ 传输。各步调用 web_download.py / pipeline.py，日志实时显示。
"""
import os
import re
import sys
import html as _html

from PySide6.QtCore import (Qt, QProcess, QProcessEnvironment, QPointF, QTimer,
                            QPropertyAnimation, QEasingCurve, QObject, QEvent)
from PySide6.QtGui import (QFont, QTextCursor, QColor, QPixmap, QPainter,
                           QPen, QPolygonF, QAction)
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QPlainTextEdit,
    QVBoxLayout, QHBoxLayout, QFrame, QCheckBox, QScrollArea, QComboBox,
    QFileDialog, QMessageBox, QGraphicsDropShadowEffect, QMenu, QProgressBar,
    QListView,
)


def _style_combo(cb):
    """给下拉框换成可样式化的列表视图，样式直接钉在视图上（弹层容器不吃全局 QSS）。"""
    v = QListView(cb)
    v.setCursor(Qt.PointingHandCursor)
    v.setStyleSheet("""
        QListView { background: #ffffff; border: 1px solid #dfe4ef;
                    border-radius: 0px; padding: 4px; outline: none;
                    color: #1f2937; }
        QListView::item { min-height: 30px; padding: 4px 10px;
                    border-radius: 6px; color: #1f2937; background: transparent; }
        QListView::item:hover { background: #f1f4fb; color: #1f2937; }
        QListView::item:selected { background: #eef2ff; color: #4f46e5; }
    """)
    cb.setView(v)
    # 弹层外框也钉白底，避免黑边
    try:
        v.window().setStyleSheet("background: #ffffff; border: 1px solid #dfe4ef;")
    except Exception:
        pass


def app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP = app_dir()
INBOX = os.path.join(APP, "_inbox")
OUTPUTS = os.path.join(APP, "_outputs")
REPORTS = os.path.join(APP, "_reports")


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


DEFAULT_URL = _endpoint("RMP_BASE")
DEFAULT_SHEET = _endpoint("QQ_SHEET_URL")


def _read_version(base) -> str:
    try:
        with open(os.path.join(base, "version.txt"), encoding="utf-8-sig") as f:
            return f.read().strip()
    except Exception:
        return "0"


VERSION = _read_version(APP)
# 热更新只同步这几个小脚本（几十 KB），不动 Chromium/PySide6 等大运行时
UPDATE_FILES = ["pipeline.py", "web_download.py", "qq_schedule.py",
                "xf_engine.py", "send_mail.py",
                "asr_pipeline_gui.py", "version.txt",
                os.path.join("tts_tools", "asr_tts_tool.py")]

QSS = """
* { font-family: 'Microsoft YaHei UI','Segoe UI'; }
QWidget#root { background: #f5f7fc; }
QLabel { color: #334155; font-size: 13px; }

#header { background: qlineargradient(x1:0,y1:0, x2:1,y2:0.9,
          stop:0 #4338ca, stop:0.45 #4f46e5, stop:1 #3b82f6); }
#hTitle { color: #ffffff; font-size: 21px; font-weight: 800; letter-spacing: 1px; }
#hSub   { color: #dbe4ff; font-size: 12px; letter-spacing: 2px; }
#verPill { background: rgba(255,255,255,0.16); color: #eef2ff;
           border: 1px solid rgba(255,255,255,0.22); border-radius: 12px;
           padding: 4px 14px; font-size: 11.5px; font-weight: 600; }
QScrollArea#scroll { background: #f5f7fc; border: none; }
QScrollBar:vertical { background: transparent; width: 10px; margin: 2px; }
QScrollBar::handle:vertical { background: #cbd3e3; border-radius: 5px; min-height: 30px; }
QScrollBar::handle:vertical:hover { background: #a9b4cb; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }

#card { background: #ffffff; border: 1px solid #e6eaf4; border-radius: 16px; }
#badge { min-width: 38px; max-width: 38px; min-height: 38px; max-height: 38px;
         border-radius: 19px; color: #ffffff; font-weight: 800; font-size: 16px;
         background: qlineargradient(x1:0,y1:0, x2:1,y2:1,
                     stop:0 #6366f1, stop:1 #4f46e5); }
#stepTitle { font-size: 16px; font-weight: 700; color: #0f172a; letter-spacing: 0.5px; }
#hint { color: #8b94a8; font-size: 12px; }
#okBadge { color: #059669; font-weight: 600; font-size: 12.5px; }
#muted   { color: #9aa3b8; font-size: 12.5px; }
#fieldLbl { color: #475569; font-size: 12.5px; font-weight: 600; }

QLineEdit { background: #f8fafc; border: 1px solid #dbe1ec; border-radius: 9px;
            padding: 8px 11px; font-size: 13px; selection-background-color: #c7d2fe; }
QLineEdit:focus { border: 1px solid #4f46e5; background: #ffffff; }
QComboBox { background: #f8fafc; border: 1px solid #dbe1ec; border-radius: 9px;
            padding: 7px 10px; font-size: 13px; }
QComboBox:focus { border: 1px solid #4f46e5; }
QComboBox::drop-down { border: none; width: 20px; }
QComboBox::down-arrow { image: none; width: 0; height: 0; margin-right: 7px;
            border-left: 4px solid transparent; border-right: 4px solid transparent;
            border-top: 5px solid #6b7280; }
QComboBox QAbstractItemView { background: #ffffff; border: 1px solid #dfe4ef;
            border-radius: 8px; outline: none; padding: 4px;
            selection-background-color: #eef2ff; selection-color: #4f46e5; }
QComboBox QAbstractItemView::item { min-height: 30px; padding: 4px 10px;
            border-radius: 6px; color: #1f2937; background: transparent; }
QComboBox QAbstractItemView::item:hover { background: #f1f4fb; color: #1f2937; }
QComboBox QAbstractItemView::item:selected { background: #eef2ff; color: #4f46e5; }

#stepper { background: #f8fafc; border: 1px solid #dbe1ec; border-radius: 9px; }
#stepBtn { background: transparent; border: none; color: #6b7280;
           font-size: 18px; font-weight: 700; }
#stepBtn:hover { background: #e9edf6; color: #4f46e5; }
QPushButton#stepBtn[side="l"] { border-top-left-radius: 8px; border-bottom-left-radius: 8px; }
QPushButton#stepBtn[side="r"] { border-top-right-radius: 8px; border-bottom-right-radius: 8px; }
#stepVal { color: #111827; font-size: 14px; font-weight: 600;
           border-left: 1px solid #e5e9f2; border-right: 1px solid #e5e9f2; }

QCheckBox { spacing: 7px; color: #374151; font-size: 13px; }
QCheckBox::indicator { width: 17px; height: 17px; border: 1px solid #c3cad8;
           border-radius: 5px; background: #ffffff; }
QCheckBox::indicator:checked { background: #4f46e5; border: 1px solid #4f46e5;
           image: url(__CHECK_ICON__); }

QPushButton { background: #eef1f7; color: #364152; border: 1px solid #e2e7f1;
           border-radius: 9px; padding: 8px 14px; font-size: 13px; font-weight: 600; }
QPushButton:hover { background: #e4e9f4; border-color: #d3dae8; }
QPushButton:pressed { background: #d9e0ef; padding-top: 9px; padding-bottom: 7px; }
QPushButton:disabled { color: #aab2c0; background: #f1f3f8; border-color: #edf0f6; }
QPushButton::menu-indicator { subcontrol-origin: padding; subcontrol-position: right center;
           width: 0; height: 0; margin-right: 10px;
           border-left: 4px solid transparent; border-right: 4px solid transparent;
           border-top: 5px solid #6b7280; }
QMenu { background: #ffffff; border: 1px solid #dfe4ef; border-radius: 10px; padding: 6px; }
QMenu::item { padding: 8px 22px; border-radius: 7px; font-size: 13px; color: #1f2937; }
QMenu::item:selected { background: #eef2ff; color: #4f46e5; }
QMenu::separator { height: 1px; background: #eef1f6; margin: 5px 8px; }
#miniGhost { background: transparent; color: #8a93a6; border: none;
           font-size: 12px; font-weight: 500; padding: 3px 6px; }
#miniGhost:hover { color: #4f46e5; }
QPushButton#primary { background: #4f46e5; color: #ffffff; border: none; }
QPushButton#primary:hover { background: #4338ca; }
QPushButton#primary:disabled { background: #b9b6ec; }
QPushButton#cta { background: #4f46e5; color: #ffffff; border: none;
           font-size: 15px; font-weight: 700; padding: 13px 30px; border-radius: 11px; }
QPushButton#cta:hover { background: #4338ca; }
QPushButton#cta:disabled { background: #b9b6ec; }
QPushButton#ghost { background: transparent; color: #4f46e5; border: none; font-weight: 600; }
QPushButton#ghost:hover { color: #4338ca; }
QPushButton#danger { background: #fef2f2; color: #dc2626; border: 1px solid #fbcaca; }
QPushButton#danger:hover { background: #fee2e2; }
QPushButton#danger:disabled { color: #e6b9b9; background: #fdf5f5; border-color: #f6e2e2; }

#logWrap { background: #0b1220; border: 1px solid #0b1220; border-radius: 12px; }
#log { background: #0b1220; color: #cbd5e1; border: none;
       font-family: 'Cascadia Mono',Consolas,'Courier New'; font-size: 12px; }
#logTitle { color: #8891a5; font-size: 12px; font-weight: 600; letter-spacing: 1px; }
QProgressBar#pbar { background: #1e293b; border: none; border-radius: 4px;
       min-height: 8px; max-height: 8px; }
QProgressBar#pbar::chunk { border-radius: 4px;
       background: qlineargradient(x1:0,y1:0, x2:1,y2:0,
                   stop:0 #6366f1, stop:0.6 #4f8bf6, stop:1 #38bdf8); }
#status { color: #6b7280; font-size: 12px; }
"""


class HoverGlow(QObject):
    """按钮悬停动画：只渐变背景色，不加阴影效果（避免文字发虚）。"""
    # objectName -> (常态背景, 悬停背景)
    COLORS = {
        "": ("#eef1f7", "#dde4f2"),
        "primary": ("#4f46e5", "#4338ca"),
        "cta": ("#4f46e5", "#4338ca"),
        "danger": ("#fef2f2", "#fee2e2"),
    }
    SKIP = {"ghost", "miniGhost", "stepBtn", "stepMinus", "stepPlus"}

    def _animate(self, btn, to_hover: bool):
        from PySide6.QtCore import QVariantAnimation
        name = btn.objectName()
        if name in self.SKIP or name not in self.COLORS:
            return
        base, hov = self.COLORS[name]
        start = QColor(hov if not to_hover else base)
        end = QColor(hov if to_hover else base)
        old = getattr(btn, "_bg_anim", None)
        if old is not None:
            try:
                if old.state() == old.State.Running:
                    start = old.currentValue()
                old.stop()
            except Exception:
                pass
        a = QVariantAnimation(btn)
        a.setDuration(150); a.setStartValue(start); a.setEndValue(end)
        a.setEasingCurve(QEasingCurve.OutCubic)
        a.valueChanged.connect(
            lambda c, b=btn: b.setStyleSheet(f"background-color: {c.name()};"))
        a.finished.connect(
            lambda b=btn, t=to_hover: (None if t else b.setStyleSheet("")))
        btn._bg_anim = a
        a.start()

    def eventFilter(self, obj, ev):
        if isinstance(obj, QPushButton) and obj.isEnabled():
            if ev.type() == QEvent.Enter:
                obj.setCursor(Qt.PointingHandCursor)
                self._animate(obj, True)
            elif ev.type() == QEvent.Leave:
                self._animate(obj, False)
        return False


class Card(QFrame):
    def __init__(self, num, title, hint):
        super().__init__()
        self.setObjectName("card")
        sh = QGraphicsDropShadowEffect(self)
        sh.setBlurRadius(22); sh.setColor(QColor(31, 41, 90, 26)); sh.setOffset(0, 5)
        self.setGraphicsEffect(sh)
        self._shadow = sh
        self._hover_anims = []

        row = QHBoxLayout(self)
        row.setContentsMargins(18, 11, 18, 11); row.setSpacing(15)

        badge = QLabel(str(num)); badge.setObjectName("badge"); badge.setAlignment(Qt.AlignCenter)
        row.addWidget(badge, 0, Qt.AlignTop)

        self.mid = QVBoxLayout(); self.mid.setSpacing(5)
        t = QLabel(title); t.setObjectName("stepTitle")
        h = QLabel(hint); h.setObjectName("hint"); h.setWordWrap(True)
        self.mid.addWidget(t); self.mid.addWidget(h)
        row.addLayout(self.mid, 1)

        self.actions = QVBoxLayout(); self.actions.setSpacing(8)
        self.actions.setAlignment(Qt.AlignVCenter)
        row.addLayout(self.actions, 0)

    def _animate_shadow(self, blur, dy):
        self._hover_anims.clear()
        for prop, val in ((b"blurRadius", blur), (b"yOffset", dy)):
            a = QPropertyAnimation(self._shadow, prop, self)
            a.setDuration(180); a.setEndValue(val)
            a.setEasingCurve(QEasingCurve.OutCubic)
            a.start()
            self._hover_anims.append(a)

    def enterEvent(self, e):
        self._animate_shadow(34, 9)     # 悬停：影子放大上浮
        super().enterEvent(e)

    def leaveEvent(self, e):
        self._animate_shadow(22, 5)     # 离开：恢复
        super().leaveEvent(e)


class Stepper(QWidget):
    """现代 [−] 值 [+] 步进器，替代古老的 QSpinBox。"""
    def __init__(self, value=5, lo=1, hi=99):
        super().__init__()
        self.lo, self.hi, self.val = lo, hi, value
        self.setObjectName("stepper")
        h = QHBoxLayout(self); h.setContentsMargins(0, 0, 0, 0); h.setSpacing(0)
        self.minus = QPushButton("−")
        self.plus = QPushButton("+")
        self.lbl = QLabel(str(value)); self.lbl.setObjectName("stepVal")
        self.lbl.setAlignment(Qt.AlignCenter); self.lbl.setFixedWidth(48)
        for b, side in ((self.minus, "l"), (self.plus, "r")):
            b.setObjectName("stepBtn"); b.setProperty("side", side)
            b.setFixedSize(34, 36); b.setFocusPolicy(Qt.NoFocus)
        self.minus.clicked.connect(lambda: self._set(self.val - 1))
        self.plus.clicked.connect(lambda: self._set(self.val + 1))
        h.addWidget(self.minus); h.addWidget(self.lbl); h.addWidget(self.plus)

    def _set(self, v):
        self.val = max(self.lo, min(self.hi, v))
        self.lbl.setText(str(self.val))

    def value(self):
        return self.val


class GUI(QWidget):
    def __init__(self):
        super().__init__()
        self.proc = None
        self.setObjectName("root")
        self.setWindowTitle("ASR 语料处理流水线")

        root = QVBoxLayout(self); root.setContentsMargins(0, 0, 0, 0); root.setSpacing(0)
        root.addWidget(self._header())

        # 内容放进滚动区：放不下时可滚动，任何分辨率都不会有东西不可见
        self.scroll = QScrollArea(); self.scroll.setObjectName("scroll")
        self.scroll.setWidgetResizable(True); self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        body = QWidget(); body.setObjectName("root")
        bl = QVBoxLayout(body); bl.setContentsMargins(18, 12, 18, 6); bl.setSpacing(10)
        self._cards = [self._step1(), self._step2(), self._step3(), self._step4()]
        for card in self._cards:
            bl.addWidget(card)
        bl.addWidget(self._utility())
        bl.addStretch(1)
        self.scroll.setWidget(body)
        root.addWidget(self.scroll, 1)

        # 日志常驻底部，始终可见（只有上方卡片区会滚动）
        logrow = QWidget(); logrow.setObjectName("root")
        lrl = QHBoxLayout(logrow); lrl.setContentsMargins(18, 0, 18, 6); lrl.setSpacing(0)
        lrl.addWidget(self._logbox())
        root.addWidget(logrow)

        self.status = QLabel("  ● 就绪"); self.status.setObjectName("status")
        self.status.setFixedHeight(24)
        root.addWidget(self.status)

        # 按当前桌面可用分辨率自适应窗口大小并居中
        scr = QApplication.primaryScreen().availableGeometry()
        w = min(980, scr.width() - 100)
        h = min(860, scr.height() - 90)
        self.setMinimumSize(min(680, w), min(440, h))
        self.resize(w, h)
        self.move(scr.x() + (scr.width() - w) // 2, scr.y() + (scr.height() - h) // 2)

        os.makedirs(INBOX, exist_ok=True)
        self.refresh_inbox()

        def _to_top():
            try:
                self.btn_login.setFocus()   # 焦点放顶部控件，避免滚动条被带下去
                self.scroll.verticalScrollBar().setValue(0)
            except Exception:
                pass
        QTimer.singleShot(0, _to_top)
        QTimer.singleShot(160, _to_top)

        # 入场动画：窗口淡入 + 卡片依次展开
        self._anims = []
        self.setWindowOpacity(0.0)
        fade = QPropertyAnimation(self, b"windowOpacity", self)
        fade.setDuration(260); fade.setStartValue(0.0); fade.setEndValue(1.0)
        fade.setEasingCurve(QEasingCurve.OutCubic)
        self._anims.append(fade)
        QTimer.singleShot(0, fade.start)
        QTimer.singleShot(60, self._animate_cards_in)

        # 运行中呼吸灯
        self._pulse = QTimer(self); self._pulse.setInterval(450)
        self._pulse.timeout.connect(self._pulse_tick)
        self._pulse_on = False

    def _animate_cards_in(self):
        for i, card in enumerate(self._cards):
            target = max(card.sizeHint().height(), 60)
            card.setMaximumHeight(0)
            anim = QPropertyAnimation(card, b"maximumHeight", self)
            anim.setDuration(300); anim.setStartValue(0); anim.setEndValue(target)
            anim.setEasingCurve(QEasingCurve.OutCubic)
            anim.finished.connect(lambda c=card: c.setMaximumHeight(16777215))
            self._anims.append(anim)
            QTimer.singleShot(90 * i, anim.start)

    def _pulse_tick(self):
        self._pulse_on = not self._pulse_on
        color = "#4f46e5" if self._pulse_on else "#a5b4fc"
        self.status.setStyleSheet(f"color:{color};font-size:12px;")

    # ── 头部 ──
    def _header(self):
        h = QFrame(); h.setObjectName("header"); h.setFixedHeight(72)
        hl = QHBoxLayout(h); hl.setContentsMargins(24, 0, 24, 0)
        v = QVBoxLayout(); v.setSpacing(2)
        v.addStretch(1)
        t = QLabel("ASR 语料处理流水线"); t.setObjectName("hTitle")
        s = QLabel("下载语料  ›  一键处理  ›  跨网段传输  ›  引擎与邮件")
        s.setObjectName("hSub")
        v.addWidget(t); v.addWidget(s); v.addStretch(1)
        hl.addLayout(v); hl.addStretch(1)
        pill = QLabel(f"v {VERSION}"); pill.setObjectName("verPill")
        hl.addWidget(pill, 0, Qt.AlignVCenter)
        return h

    # ── 步骤1：下载 ──
    def _step1(self):
        c = Card("1", "下载语料", "输入需求管理平台网址 → 登录一次 → 一键下载全部语料到 _inbox")
        row = QHBoxLayout(); row.setSpacing(9)
        lbl = QLabel("网址"); lbl.setObjectName("fieldLbl")
        self.url_edit = QLineEdit(DEFAULT_URL)
        self.url_edit.setPlaceholderText("需求管理平台网址")
        self.url_edit.setMinimumWidth(160)
        self.url_edit.setFocusPolicy(Qt.ClickFocus)   # 避免启动即抢焦点弹输入法
        row.addWidget(lbl); row.addWidget(self.url_edit, 1)
        self.mid_add(c, row)

        # 按排期表（腾讯文档）过滤下载：可指定 tab / 责任人
        row2 = QHBoxLayout(); row2.setSpacing(9)
        lbl2 = QLabel("排期表"); lbl2.setObjectName("fieldLbl")
        self.sheet_edit = QLineEdit(DEFAULT_SHEET)
        self.sheet_edit.setMinimumWidth(90)
        self.sheet_edit.setFocusPolicy(Qt.ClickFocus)
        lbl3 = QLabel("tab"); lbl3.setObjectName("fieldLbl")
        self.tab_combo = QComboBox(); self.tab_combo.setFixedWidth(150)
        self.tab_combo.setEditable(True)          # 可选可手输
        self.tab_combo.addItems(["（默认tab）", "2026-06需求单", "2026-07需求单",
                                 "2026实体更新排期表", "202605问题单",
                                 "语种定制人员情况", "版本数量统计"])
        _style_combo(self.tab_combo)
        self.btn_tabs_refresh = QPushButton("↻"); self.btn_tabs_refresh.setObjectName("miniGhost")
        self.btn_tabs_refresh.setToolTip("从排期表刷新 tab 列表")
        self.btn_tabs_refresh.setFixedWidth(26)
        self.btn_tabs_refresh.clicked.connect(self.refresh_tabs)
        lbl4 = QLabel("责任人"); lbl4.setObjectName("fieldLbl")
        self.user_edit = QLineEdit(); self.user_edit.setFixedWidth(80)
        self.user_edit.setPlaceholderText("可留空")
        self.btn_sheet_view = QPushButton("查看"); self.btn_sheet_view.setObjectName("ghost")
        self.btn_sheet_view.clicked.connect(lambda: self.do_sheet(False))
        _iconize(self.btn_sheet_view, 0xE721, "#4f46e5")      # 搜索
        self.btn_sheet_dl = QPushButton("按排期下载"); self.btn_sheet_dl.setObjectName("primary")
        self.btn_sheet_dl.clicked.connect(lambda: self.do_sheet(True))
        _iconize(self.btn_sheet_dl, 0xE896, "#ffffff")
        for w in (lbl2, self.sheet_edit, lbl3, self.tab_combo, self.btn_tabs_refresh,
                  lbl4, self.user_edit, self.btn_sheet_view, self.btn_sheet_dl):
            row2.addWidget(w, 1 if w is self.sheet_edit else 0)
        self.mid_add(c, row2)

        self.inbox_lbl = QLabel(""); self.inbox_lbl.setObjectName("muted")
        c.mid.addWidget(self.inbox_lbl)

        self.btn_login = QPushButton("登录"); self.btn_login.setObjectName("primary")
        self.btn_login.setMinimumWidth(76); self.btn_login.clicked.connect(self.do_login)
        _iconize(self.btn_login, 0xE77B, "#ffffff")           # 人像
        self.scope = QComboBox(); self.scope.addItems(["全部", "只本地", "只云端"])
        self.scope.setFixedWidth(84)
        _style_combo(self.scope)
        self.btn_dl = QPushButton("下载"); self.btn_dl.setObjectName("primary")
        self.btn_dl.clicked.connect(self.do_download)
        _iconize(self.btn_dl, 0xE896, "#ffffff")              # 下载
        btns = QHBoxLayout(); btns.setSpacing(8)
        btns.addWidget(self.btn_login); btns.addWidget(self.scope); btns.addWidget(self.btn_dl)
        c.actions.addLayout(btns)
        return c

    # ── 步骤2：处理 ──
    def _step2(self):
        c = Card("2", "处理语料", "自动：备份 → 按排期归档 → 清洗标准化 → 生成 TTS 测试集 → 打包 → 核验")
        row = QHBoxLayout(); row.setSpacing(9)
        lbl = QLabel("产物保留份数"); lbl.setObjectName("fieldLbl")
        self.keep = Stepper(5, 1, 99)
        self.tts_cb = QCheckBox("生成 TTS 测试集"); self.tts_cb.setChecked(True)
        row.addWidget(lbl); row.addWidget(self.keep)
        row.addSpacing(18); row.addWidget(self.tts_cb); row.addStretch(1)
        self.mid_add(c, row)
        self.btn_run = QPushButton("开始处理"); self.btn_run.setObjectName("cta")
        self.btn_run.clicked.connect(self.do_run)
        _iconize(self.btn_run, 0xE768, "#ffffff", 18)         # 播放
        c.actions.addWidget(self.btn_run)
        return c

    # ── 步骤3：传输 ──
    def _step3(self):
        c = Card("3", "传输产物", "把本批 new_corpus + TTS 测试集上传到接收网段（rftctl）")
        row = QHBoxLayout(); row.setSpacing(9)
        self.auto_send = QCheckBox("处理完自动传输"); self.auto_send.setChecked(True)
        lbl = QLabel("接收网段"); lbl.setObjectName("fieldLbl")
        self.recv = QLineEdit("rdg,dtn"); self.recv.setFixedWidth(130)
        row.addWidget(self.auto_send); row.addSpacing(14)
        row.addWidget(lbl); row.addWidget(self.recv); row.addStretch(1)
        self.mid_add(c, row)
        self.btn_send = QPushButton("重发最近产物"); self.btn_send.setObjectName("ghost")
        self.btn_send.clicked.connect(self.do_send)
        _iconize(self.btn_send, 0xE72C, "#4f46e5")            # 刷新
        c.actions.addWidget(self.btn_send)
        return c

    # ── 步骤4：引擎与邮件 ──
    def _step4(self):
        c = Card(4, "引擎与邮件", "训完模型后：查最新上架引擎（支持模糊搜索：语种/车厂/单号/关键词）→ 下载/传内网 → 发上架通知")
        row = QHBoxLayout(); row.setSpacing(9)
        lbl1 = QLabel("语种"); lbl1.setObjectName("fieldLbl")
        self.eng_lang = QLineEdit(); self.eng_lang.setFixedWidth(110)
        self.eng_lang.setPlaceholderText("泰语 / 单号")
        lbl2 = QLabel("车厂"); lbl2.setObjectName("fieldLbl")
        self.eng_brand = QLineEdit(); self.eng_brand.setFixedWidth(90)
        self.eng_brand.setPlaceholderText("极氪")
        b_list = QPushButton("查引擎"); b_list.clicked.connect(lambda: self.do_engine("list"))
        _iconize(b_list, 0xE721)                              # 搜索
        b_dl = QPushButton("下载到本机"); b_dl.clicked.connect(lambda: self.do_engine("local"))
        _iconize(b_dl, 0xE896)                                # 下载
        b_rdg = QPushButton("传rdg"); b_rdg.clicked.connect(lambda: self.do_engine("rdg"))
        _iconize(b_rdg, 0xE898)                               # 上传
        for w in (lbl1, self.eng_lang, lbl2, self.eng_brand, b_list, b_dl, b_rdg):
            row.addWidget(w)
        row.addStretch(1)
        self.mid_add(c, row)

        # 相关登录/配置（小号入口，归位到本卡片）
        acc = QHBoxLayout(); acc.setSpacing(2)
        hint = QLabel("账号：")
        hint.setObjectName("hint")
        acc.addWidget(hint)
        for txt, fn in [("引擎登录", lambda: self.run(["engine", "login"])),
                        ("登录E3", lambda: self.run(["engine", "login-devops"])),
                        ("邮件配置", self.open_mail_config)]:
            b = QPushButton(txt); b.setObjectName("miniGhost"); b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(fn); acc.addWidget(b)
        acc.addStretch(1)
        self.mid_add(c, acc)

        mailbox = QVBoxLayout(); mailbox.setSpacing(6)
        self.btn_mail_prev = QPushButton("预览邮件"); self.btn_mail_prev.setObjectName("ghost")
        self.btn_mail_prev.clicked.connect(lambda: self.do_mail(dry=True))
        _iconize(self.btn_mail_prev, 0xE8A5, "#4f46e5")       # 文档预览
        self.btn_mail_send = QPushButton("发送上架邮件"); self.btn_mail_send.setObjectName("primary")
        self.btn_mail_send.clicked.connect(lambda: self.do_mail(dry=False))
        _iconize(self.btn_mail_send, 0xE715, "#ffffff")       # 邮件
        mailbox.addWidget(self.btn_mail_prev); mailbox.addWidget(self.btn_mail_send)
        c.actions.addLayout(mailbox)
        return c

    def mid_add(self, card, layout):
        w = QWidget(); w.setStyleSheet("background:transparent;")
        w.setLayout(layout)
        card.mid.addWidget(w)

    # ── 工具行（分组菜单，清爽）──
    def _utility(self):
        w = QFrame(); w.setStyleSheet("background:transparent;")
        h = QHBoxLayout(w); h.setContentsMargins(2, 0, 2, 0); h.setSpacing(8)

        b_add = QPushButton("添加语料"); b_add.clicked.connect(self.add_files)
        _iconize(b_add, 0xE710)                               # 加号
        h.addWidget(b_add)

        b_open = QPushButton("打开  ")
        _iconize(b_open, 0xE8B7)                              # 文件夹
        m_open = QMenu(b_open)
        for txt, fn in [("排期表", self.open_schedule),
                        ("_inbox（待处理语料）", lambda: self.open_dir(INBOX)),
                        ("_outputs（产物）", lambda: self.open_dir(OUTPUTS)),
                        ("_reports（报告）", lambda: self.open_dir(REPORTS)),
                        ("_engines（引擎）", lambda: self.open_dir(os.path.join(APP, "_engines")))]:
            act = QAction(txt, m_open); act.triggered.connect(fn); m_open.addAction(act)
        b_open.setMenu(m_open)
        h.addWidget(b_open)

        b_tool = QPushButton("工具  ")
        _iconize(b_tool, 0xE713)                              # 齿轮
        m_tool = QMenu(b_tool)
        for txt, fn in [("只核验", lambda: self.run(["verify"])),
                        ("排期状态", lambda: self.run(["status"])),
                        ("检查更新", self.check_update)]:
            act = QAction(txt, m_tool); act.triggered.connect(fn); m_tool.addAction(act)
        b_tool.setMenu(m_tool)
        h.addWidget(b_tool)

        h.addStretch(1)
        self.btn_stop = QPushButton("停止"); self.btn_stop.setObjectName("danger")
        self.btn_stop.clicked.connect(self.stop); self.btn_stop.setEnabled(False)
        _iconize(self.btn_stop, 0xE71A, "#dc2626")            # 停止
        h.addWidget(self.btn_stop)
        return w

    # ── 日志 ──
    def _logbox(self):
        wrap = QFrame(); wrap.setObjectName("logWrap")
        v = QVBoxLayout(wrap); v.setContentsMargins(12, 10, 12, 12); v.setSpacing(6)
        head = QHBoxLayout(); head.setSpacing(10)
        lab = QLabel("运行日志"); lab.setObjectName("logTitle")
        self.pbar = QProgressBar(); self.pbar.setObjectName("pbar")
        self.pbar.setTextVisible(False); self.pbar.setRange(0, 100)
        self.pbar.setValue(0); self.pbar.hide()
        self.plabel = QLabel(""); self.plabel.setObjectName("logTitle"); self.plabel.hide()
        clr = QPushButton("清空"); clr.setObjectName("ghost"); clr.clicked.connect(lambda: self.log.clear())
        head.addWidget(lab); head.addWidget(self.pbar, 1)
        head.addWidget(self.plabel); head.addWidget(clr)
        v.addLayout(head)
        self.log = QPlainTextEdit(); self.log.setObjectName("log"); self.log.setReadOnly(True)
        self.log.setMinimumHeight(84)
        v.addWidget(self.log, 1)
        # 无真实进度时的“流动”动画
        self._busy_val = 0
        self._busy = QTimer(self); self._busy.setInterval(28)
        self._busy.timeout.connect(self._busy_tick)
        self._prog_real = False
        self._line_buf = ""
        return wrap

    def _busy_tick(self):
        self._busy_val = (self._busy_val + 2) % 100
        self.pbar.setValue(self._busy_val)

    # ── inbox 计数 ──
    def refresh_inbox(self):
        try:
            n = len([f for f in os.listdir(INBOX)
                     if f.lower().endswith(".xlsx") and not f.startswith("~$")])
        except Exception:
            n = 0
        if n:
            self.inbox_lbl.setText(f"✓  _inbox 已有 {n} 个待处理语料")
            self.inbox_lbl.setStyleSheet("color:#059669; font-weight:600; font-size:12.5px;")
        else:
            self.inbox_lbl.setText("○  _inbox 为空 —— 先下载或添加语料")
            self.inbox_lbl.setStyleSheet("color:#9aa3b2; font-size:12.5px;")

    # ── 日志输出（按语义着色）──
    _C_OK   = "#4ade80"
    _C_ERR  = "#f87171"
    _C_WARN = "#fbbf24"
    _C_CMD  = "#7dd3fc"
    _C_DIM  = "#64748b"
    _C_TXT  = "#cbd5e1"

    _C_PROG = "#38bdf8"
    # “失败/错误/跳过”后面跟 空列表/0 → 不算坏消息
    _RE_EMPTY_BAD = re.compile(r'(失败|错误)[/，,、\w]*[:：]?\s*(\[\]|0\b|0\s*个)')
    _RE_EMPTY_SKIP = re.compile(r'跳过[:：]?\s*(\[\]|0\b|0\s*个)')

    def _line_color(self, ln: str) -> str:
        s = ln.strip()
        if not s:
            return self._C_TXT
        if s.startswith("$"):
            return self._C_CMD
        if re.match(r'^\[\d+/\d+\]', s):          # 进度行
            return self._C_PROG
        if "退出码" in s:                          # 结束行按退出码定色
            return self._C_OK if re.search(r'退出码\s*0\b', s) else self._C_WARN
        has_bad = any(k in s for k in ("✗", "失败", "错误", "FAIL", "[X]", "Traceback", "Error"))
        if has_bad and not self._RE_EMPTY_BAD.search(s):
            return self._C_ERR
        if any(k in s for k in ("✓", "成功", "完成", "PASS", "已发送", "已下载", "已保存", "已生成", "已刷新")):
            return self._C_OK
        has_warn = any(k in s for k in ("警告", "warn", "跳过", "未匹配", "缺少", "待确认"))
        if has_warn and not self._RE_EMPTY_SKIP.search(s):
            return self._C_WARN
        if set(s) <= set("─═-= "):
            return self._C_DIM
        return self._C_TXT

    def append(self, text):
        """着色输出：按行缓冲（QProcess 是分块到达的），整行落色。"""
        self._line_buf += text
        *lines, self._line_buf = self._line_buf.split("\n")
        for ln in lines:
            color = self._line_color(ln)
            esc = _html.escape(ln) or "&nbsp;"
            bold = " font-weight:600;" if color in (self._C_CMD, self._C_ERR) else ""
            self.log.appendHtml(
                f'<span style="color:{color};{bold} white-space:pre;">{esc}</span>')
        self.log.moveCursor(QTextCursor.End)

    def _flush_log(self):
        if self._line_buf:
            self.append("\n")

    # ── 进度条驱动 ──
    _RE_IN = re.compile(r'\[(\d+)/(\d+)\]')
    _RE_PCT = re.compile(r'(\d{1,3})(?:\.\d+)?%')

    def _feed_progress(self, chunk: str):
        m = None
        for m in self._RE_IN.finditer(chunk):
            pass
        if m:
            cur, total = int(m.group(1)), int(m.group(2))
            if total > 0:
                self._prog_real = True
                self._busy.stop()
                self.pbar.setRange(0, total); self.pbar.setValue(cur)
                self.plabel.setText(f"{cur}/{total}"); self.plabel.show()
                return
        m2 = None
        for m2 in self._RE_PCT.finditer(chunk):
            pass
        if m2 and not self._prog_real:
            pct = min(100, int(m2.group(1)))
            self._busy.stop()
            self.pbar.setRange(0, 100); self.pbar.setValue(pct)
            self.plabel.setText(f"{pct}%"); self.plabel.show()

    def build_cmd(self, tokens):
        if getattr(sys, "frozen", False):
            return sys.executable, tokens
        if tokens and tokens[0] == "sheet":
            return sys.executable, [os.path.join(APP, "qq_schedule.py")] + tokens[1:]
        if tokens and tokens[0] == "engine":
            return sys.executable, [os.path.join(APP, "xf_engine.py")] + tokens[1:]
        if tokens and tokens[0] == "mail":
            return sys.executable, [os.path.join(APP, "send_mail.py")] + tokens[1:]
        if tokens and tokens[0] in ("login", "download"):
            return sys.executable, [os.path.join(APP, "web_download.py")] + tokens
        return sys.executable, [os.path.join(APP, "pipeline.py")] + tokens

    def busy(self):
        return self.proc is not None and self.proc.state() != QProcess.NotRunning

    def run(self, tokens):
        if self.busy():
            QMessageBox.warning(self, "忙", "已有任务在运行，请等它结束或先停止。")
            return
        prog, args = self.build_cmd(tokens)
        self.append(f"\n$ {' '.join(tokens)}\n" + "─" * 58 + "\n")
        # 进度条：先流动动画，解析到 [i/N] 或 % 后切真实进度
        self._prog_real = False
        self._busy_val = 0
        self.pbar.setRange(0, 100); self.pbar.setValue(0)
        self.pbar.show(); self.plabel.hide()
        self._busy.start()
        self.set_running(True)
        self.proc = QProcess(self)
        self.proc.setWorkingDirectory(APP)
        self.proc.setProcessChannelMode(QProcess.MergedChannels)
        env = QProcessEnvironment.systemEnvironment()
        env.insert("PYTHONIOENCODING", "utf-8")
        self.proc.setProcessEnvironment(env)
        self.proc.readyReadStandardOutput.connect(self._on_out)
        self.proc.finished.connect(self._on_done)
        self.proc.errorOccurred.connect(lambda e: self.append(f"[进程错误] {e}\n"))
        self.proc.start(prog, args)

    def _on_out(self):
        data = bytes(self.proc.readAllStandardOutput()).decode("utf-8", "replace")
        self.append(data)
        self._feed_progress(data)
        if getattr(self, "_tab_capture", None) is not None:
            self._tab_capture.append(data)

    def _on_done(self, code, _status):
        self._flush_log()
        self.append("─" * 58 + f"\n[完成] 退出码 {code}\n")
        # 进度条收尾：满格后淡出
        self._busy.stop()
        self.pbar.setRange(0, 100); self.pbar.setValue(100)
        self.plabel.setText("完成" if code == 0 else "结束")
        QTimer.singleShot(1200, self.pbar.hide)
        QTimer.singleShot(1200, self.plabel.hide)
        self.set_running(False, code)
        self.refresh_inbox()
        # 「↻ 刷新tab列表」的结果回填下拉
        if getattr(self, "_tab_capture", None) is not None:
            text = "".join(self._tab_capture)
            self._tab_capture = None
            names = re.findall(r'^\s{2}\S+\s{2}(\S+)\s*$', text, re.M)
            if names:
                cur = self.tab_combo.currentText()
                self.tab_combo.clear()
                self.tab_combo.addItems(["（默认tab）"] + names)
                if cur in names:
                    self.tab_combo.setCurrentText(cur)
                self.append(f"[已刷新 tab 列表：{len(names)} 个]\n")

    def set_running(self, running, code=None):
        for b in (self.btn_run, self.btn_login, self.btn_dl, self.btn_send,
                  self.btn_sheet_view, self.btn_sheet_dl,
                  self.btn_mail_prev, self.btn_mail_send):
            b.setEnabled(not running)
        self.btn_stop.setEnabled(running)
        if running:
            self.status.setText("● 运行中…"); self.status.setStyleSheet("color:#4f46e5;font-size:12px;")
            self._pulse.start()
        else:
            self._pulse.stop()
            ok = (code == 0)
            self.status.setText("● 完成" if ok else f"● 结束（退出码 {code}）")
            self.status.setStyleSheet(f"color:{'#059669' if ok else '#b45309'};font-size:12px;")

    def stop(self):
        if self.busy():
            self.proc.kill()
            self.append("\n[已停止]\n")

    # ── 各操作 ──
    def do_login(self):
        base = self.url_edit.text().strip() or DEFAULT_URL
        self.append("提示：稍后会弹出浏览器，请登录后【关闭浏览器窗口】。\n")
        self.run(["login", "--base", base])

    def do_download(self):
        base = self.url_edit.text().strip() or DEFAULT_URL
        scope = {"全部": "all", "只本地": "local", "只云端": "cloud"}[self.scope.currentText()]
        self.run(["download", "--base", base, "--scope", scope])

    def do_engine(self, action):
        lang = self.eng_lang.text().strip()
        brand = self.eng_brand.text().strip()
        if not lang and not brand:
            QMessageBox.information(self, "引擎", "请先填 语种 或 车厂（至少一个）。")
            return
        base = ["engine"]
        args = []
        if lang:
            args += ["--lang", lang]
        if brand:
            args += ["--brand", brand]
        if action == "list":
            self.run(base + ["list"] + args)
        elif action == "local":
            self.run(base + ["get"] + args + ["--local"])
        elif action == "rdg":
            self.run(base + ["get"] + args + ["--to", "rdg"])

    def do_mail(self, dry):
        lang = self.eng_lang.text().strip()
        brand = self.eng_brand.text().strip()
        if not lang and not brand:
            QMessageBox.information(self, "邮件", "请先在「引擎与邮件」填 语种/车厂。")
            return
        args = ["mail", "engine"]
        if lang:
            args += ["--lang", lang]
        if brand:
            args += ["--brand", brand]
        if dry:
            args.append("--dry")
        elif QMessageBox.question(self, "发送确认",
                                  f"确认发送「{brand}{lang}」引擎上架邮件？\n"
                                  "（收件人在 mail_config.json 配置）") != QMessageBox.Yes:
            return
        self.run(args)

    def refresh_tabs(self):
        """跑 sheet tabs 并把结果填进下拉。"""
        self._tab_capture = []
        self.run(["sheet", "tabs", "--url", self.sheet_edit.text().strip() or DEFAULT_SHEET])

    def do_sheet(self, download):
        args = ["sheet", "read", "--url", self.sheet_edit.text().strip() or DEFAULT_SHEET]
        tab = self.tab_combo.currentText().strip()
        if tab and tab != "（默认tab）":
            args += ["--tab", tab]
        user = self.user_edit.text().strip()
        if user:
            args += ["--user", user]
        scope = {"全部": "all", "只本地": "local", "只云端": "cloud"}[self.scope.currentText()]
        args += ["--scope", scope]
        if download:
            args.append("--download")
        self.run(args)

    def do_run(self):
        if not self._inbox_has():
            if QMessageBox.question(self, "_inbox 为空",
                                    "_inbox 里没有 xlsx，仍要运行吗？") != QMessageBox.Yes:
                return
        args = ["run"]
        if not self.tts_cb.isChecked():
            args.append("--no-tts")
        if self.auto_send.isChecked():
            args += ["--send", "--receive", self.recv.text().strip() or "rdg,dtn"]
        args += ["--keep", str(self.keep.value())]
        self.run(args)

    def do_send(self):
        self.run(["send", "--receive", self.recv.text().strip() or "rdg,dtn"])

    def _inbox_has(self):
        try:
            return any(f.lower().endswith(".xlsx") and not f.startswith("~$")
                       for f in os.listdir(INBOX))
        except Exception:
            return False

    def open_dir(self, path):
        os.makedirs(path, exist_ok=True)
        try:
            os.startfile(path)
        except AttributeError:
            import subprocess
            subprocess.Popen(["xdg-open", path])

    def open_mail_config(self):
        path = os.path.join(APP, "mail_config.json")
        if not os.path.exists(path):
            self.run(["mail", "init"])
            QMessageBox.information(self, "邮件配置",
                                    "已生成 mail_config.json 模板，\n请填好邮箱账号/密码/收件人后保存。")
        try:
            os.startfile(path)
        except Exception:
            pass

    def open_schedule(self):
        path = os.path.join(APP, "排期.xlsx")
        if not os.path.exists(path):
            QMessageBox.information(self, "排期表", "排期表还不存在。\n下载语料时会自动生成，或先手动放一个 排期.xlsx。")
            return
        try:
            os.startfile(path)
        except AttributeError:
            import subprocess
            subprocess.Popen(["xdg-open", path])

    # ── 热更新（只同步小脚本，不重发大运行时）──
    def _update_token(self):
        """私有仓库鉴权 token：环境变量 ASR_UPDATE_TOKEN > 程序目录 update_token.txt。"""
        t = os.environ.get("ASR_UPDATE_TOKEN", "").strip()
        if not t:
            tp = os.path.join(APP, "update_token.txt")
            if os.path.exists(tp):
                try:
                    t = open(tp, encoding="utf-8-sig").read().strip()  # utf-8-sig 自动去 BOM
                except Exception:
                    t = ""
        return t

    def check_update(self):
        src = os.environ.get("ASR_UPDATE_SOURCE", "").strip()
        if not src:
            sp = os.path.join(APP, "update_source.txt")
            if os.path.exists(sp):
                try:
                    src = open(sp, encoding="utf-8-sig").read().strip()  # utf-8-sig 自动去 BOM
                except Exception:
                    src = ""
        if not src:
            QMessageBox.information(
                self, "检查更新",
                "未配置更新源。\n\n在程序目录建 update_source.txt，写入其一：\n"
                "· gitee raw 基地址（内网推荐）：\n"
                "  https://gitee.com/<用户>/<仓库>/raw/main\n"
                "· 网络共享文件夹（如 \\\\服务器\\share\\asr_update）\n"
                "· http(s) 更新包地址（.zip，如 GitHub 分支 zip）\n\n"
                "更新源里放最新的 pipeline.py / web_download.py /\n"
                "asr_pipeline_gui.py / version.txt 即可。")
            return
        self.append(f"[更新] 更新源：{src}\n")
        try:
            token = self._update_token()
            if src.lower().startswith("http"):
                if src.lower().rstrip("/").endswith(".zip") or "zipball" in src.lower():
                    updir = self._fetch_http_update(src, token)      # zip 包（GitHub/内网服务器）
                else:
                    updir = self._fetch_raw_base(src, token)         # 逐文件（gitee raw 等）
            else:
                updir = src                                          # 本地/共享文件夹
            if not updir or not os.path.isdir(updir):
                QMessageBox.warning(self, "检查更新", "无法访问更新源，请检查路径/网络。")
                return
            remote_v = _read_version(updir)
            if remote_v == "0":
                QMessageBox.warning(self, "检查更新", "更新源缺少 version.txt。")
                return
            if remote_v == VERSION:
                QMessageBox.information(self, "检查更新", f"已是最新版本（{VERSION}）。")
                return
            import shutil
            n = 0
            for rel in UPDATE_FILES:
                s = os.path.join(updir, rel)
                if os.path.exists(s):
                    d = os.path.join(APP, rel)
                    os.makedirs(os.path.dirname(d) or ".", exist_ok=True)
                    shutil.copy2(s, d); n += 1
                    self.append(f"[更新] {rel}\n")
            QMessageBox.information(
                self, "更新完成",
                f"已从 {VERSION} 更新到 {remote_v}（{n} 个文件）。\n请关闭并重新打开程序生效。")
        except Exception as e:
            self.append(f"[更新] 失败：{e}\n")
            QMessageBox.warning(self, "检查更新", f"更新失败：{e}")

    def _fetch_raw_base(self, base, token=None):
        """逐文件从 raw 基地址下载 UPDATE_FILES（gitee 等禁止匿名归档下载时用）。
        base 形如 https://gitee.com/<用户>/<仓库>/raw/main 。"""
        import tempfile, urllib.request, shutil
        tmp = os.path.join(tempfile.gettempdir(), "_asr_update")
        shutil.rmtree(tmp, ignore_errors=True)
        os.makedirs(tmp, exist_ok=True)
        base = base.rstrip("/")
        got_version = False
        for rel in UPDATE_FILES:
            url = base + "/" + rel.replace("\\", "/")
            dst = os.path.join(tmp, rel)
            os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
            try:
                req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
                if token:
                    req.add_header("Authorization", "Bearer " + token)
                with urllib.request.urlopen(req, timeout=60) as r:
                    data = r.read()
                with open(dst, "wb") as f:
                    f.write(data)
                if rel == "version.txt":
                    got_version = True
            except Exception as e:
                self.append(f"[更新] 跳过 {rel}：{e}\n")
        return tmp if got_version else None

    def _fetch_http_update(self, url, token=None):
        import tempfile, zipfile, urllib.request, shutil
        tmp = os.path.join(tempfile.gettempdir(), "_asr_update")
        shutil.rmtree(tmp, ignore_errors=True)
        os.makedirs(tmp, exist_ok=True)
        zp = os.path.join(tmp, "update.zip")
        req = urllib.request.Request(url)
        if token:                                   # 私有库鉴权（GitHub API zipball）
            req.add_header("Authorization", "Bearer " + token)
            req.add_header("Accept", "application/vnd.github+json")
            req.add_header("X-GitHub-Api-Version", "2022-11-28")
        with urllib.request.urlopen(req, timeout=60) as r, open(zp, "wb") as f:
            shutil.copyfileobj(r, f)
        with zipfile.ZipFile(zp) as z:
            z.extractall(tmp)
        os.remove(zp)
        entries = [os.path.join(tmp, e) for e in os.listdir(tmp)]
        dirs = [e for e in entries if os.path.isdir(e)]
        files = [e for e in entries if os.path.isfile(e)]
        if len(dirs) == 1 and not files:
            return dirs[0]
        return tmp

    def add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "选择 xlsx 语料", "", "Excel (*.xlsx)")
        if not paths:
            return
        import shutil
        os.makedirs(INBOX, exist_ok=True)
        n = 0
        for p in paths:
            try:
                shutil.copy2(p, os.path.join(INBOX, os.path.basename(p))); n += 1
            except Exception as e:
                self.append(f"[复制失败] {p}: {e}\n")
        self.append(f"[已添加 {n} 个文件到 _inbox]\n")
        self.refresh_inbox()


def _fluent_icon(cp: int, color: str = "#475569", px: int = 18):
    """用 Windows 自带 Segoe Fluent Icons 字体渲染矢量图标（Win10 回退 MDL2）。"""
    from PySide6.QtGui import QIcon, QFontDatabase
    fams = QFontDatabase.families()
    fam = ("Segoe Fluent Icons" if "Segoe Fluent Icons" in fams
           else "Segoe MDL2 Assets" if "Segoe MDL2 Assets" in fams else "")
    if not fam:
        return QIcon()
    pm = QPixmap(px * 2, px * 2)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    f = QFont(fam)
    f.setPixelSize(int(px * 1.65))
    p.setFont(f)
    p.setPen(QColor(color))
    p.drawText(pm.rect(), Qt.AlignCenter, chr(cp))
    p.end()
    return QIcon(pm)


def _iconize(btn, cp: int, color: str = "#475569", px: int = 16):
    from PySide6.QtCore import QSize
    ic = _fluent_icon(cp, color, px)
    if not ic.isNull():
        btn.setIcon(ic)
        btn.setIconSize(QSize(px, px))


def _check_icon() -> str:
    """生成白色对勾 PNG，供选中态复选框使用；返回正斜杠路径。"""
    import tempfile
    size = 16
    pm = QPixmap(size, size); pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    pen = QPen(QColor("#ffffff")); pen.setWidthF(2.1)
    pen.setCapStyle(Qt.RoundCap); pen.setJoinStyle(Qt.RoundJoin)
    p.setPen(pen)
    p.drawPolyline(QPolygonF([QPointF(3.5, 8.4), QPointF(6.6, 11.4), QPointF(12.3, 4.6)]))
    p.end()
    path = os.path.join(tempfile.gettempdir(), "_asr_check16.png")
    pm.save(path)
    return path.replace(os.sep, "/")


def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(QSS.replace("__CHECK_ICON__", _check_icon()))
    app.setFont(QFont("Microsoft YaHei UI", 10))
    glow = HoverGlow(app)
    app.installEventFilter(glow)     # 按钮悬停辉光
    w = GUI()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
