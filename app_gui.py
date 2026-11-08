# -*- coding: utf-8 -*-
"""
影子科技 · 养殖数据批量下载器（统一界面）
================================================================
一个窗口完成六种下载：
  · 母猪综合查询：日期字段 + 日期范围（超限自动按日期拆分）
  · 猪只转移    ：所在场 + 转移年月（逐条导出详情，生成清单与汇总表）
  · 母猪档案    ：阶段 + 所在场（批量下载数据，补充场 / 阶段 / 条数）
  · 遗传性状查询：基因组选种平台 · 生长性状，测定猪场 + 结测日期 + 国标编码（超 400 个分批后合并）
  · 育种值查询  ：基因组选种平台 · 按个体，只按国标编码（每批 400 个，超出分批后合并）
  · 系谱查询    ：基因组选种平台 · 系谱信息，国标编码 × 场列表 × 在场 / 离场循环查询，查全后合并
各种任务可放进同一个任务列表一次跑完：全程只开一个浏览器、只登录一次，
一个功能做完回到首页直接进入下一个功能。

命令行自动模式（批处理用，载入任务后自动开始）：
    影子科技数据下载器.exe --report "配种日期,2026-10-01,2026-10-08"
                          --transfer "信宜场,2026-06,2026-07"      （年月或日期）
                          --sow "龙马一场,1;龙马一场,2"             （1=后备/生产 2=仔猪/保育/待售）
                          --genetic "南沙母猪场16线,2026-08,2026-09"  （测定猪场,结测开始,结束；都可留空 = 不限）
                          [--genetic-codes 国标编码.txt]            （txt / csv / xlsx，用于所有遗传性状任务）
                          --ebv-codes 国标编码.txt                  （育种值查询：按这些国标编码导出）
                          --ped-codes 国标编码.txt --ped-farms "龙马一场;龙马二场" [--ped-status 在场|离场]
                                                                    （系谱查询；不写状态 = 先在场、再离场）
                          [--headless]
    母猪综合查询可写第 4 项指定公司 / 场："配种日期,2026-10-01,2026-10-08,广西扬翔农牧有限责任公司"
    影子科技数据下载器.exe --run-schedule      按「定时下载」里保存的计划运行一次，跑完自动退出
                                               （Windows 计划任务就是这样启动本程序的）
"""

import os
import re
import sys
import json
import time
import queue
import threading
import datetime
import traceback
import subprocess

import yz_core as core
import schedule_plan as sp
import ui_icons
import task_report as rp
import task_transfer as tr
import task_sowfile as sw
import task_genetic as gn
import task_ebv as ev
import task_pedigree as pd

try:
    import tkinter as tk
    import tkinter.font as tkfont
    from tkinter import ttk, filedialog, messagebox, scrolledtext
except Exception:
    raise SystemExit("当前环境缺少 tkinter，无法启动图形界面。")

APP_TITLE = "影子科技 · 养殖数据批量下载器"
KINDS = {"report": (rp.TITLE, rp), "transfer": (tr.TITLE, tr), "sow": (sw.TITLE, sw), "genetic": (gn.TITLE, gn),
         "ebv": (ev.TITLE, ev), "pedigree": (pd.TITLE, pd)}
KIND_ORDER = ["report", "transfer", "sow", "genetic", "ebv", "pedigree"]
ALL_FARMS = "（场列表中全部场）"
LIST_REPORT = rp.TITLE             # 母猪综合查询的公司 / 场列表
LIST_TRANSFER = tr.TITLE           # 猪只转移的场列表；母猪档案每个阶段各有一份（以阶段名为键）
LIST_GENETIC = gn.TITLE            # 遗传性状查询的测定猪场列表
LIST_PEDIGREE = pd.TITLE           # 系谱查询要逐个查询的场列表
G_ANY = "（不限：不选测定猪场）"
DEFAULT_TRANSFER_FARMS = ["金沙一场", "信宜场"]
MONTHS = [f"{m}月" for m in range(1, 13)]
APP_DIR_NAME = "影子科技数据下载"


def default_root():
    """还没设置过保存目录时用的默认位置：
       1) .env / 环境变量里的 YINGZI_DOWNLOAD_DIR（统一分发给多台电脑时可以预先写好）；
       2) 电脑有 D 盘 → D:\\影子科技数据下载；
       3) 否则 → 当前用户的「下载」文件夹\\影子科技数据下载。"""
    v = core.env_value("YINGZI_DOWNLOAD_DIR")
    if v:
        return os.path.normpath(os.path.expandvars(os.path.expanduser(v)))
    if os.name == "nt" and os.path.isdir("D:\\"):
        return os.path.join("D:\\", APP_DIR_NAME)
    return os.path.join(os.path.expanduser("~"), "Downloads", APP_DIR_NAME)


def dir_usable(path):
    """目录所在的盘存在、并且目录能创建 / 能写入。"""
    path = (path or "").strip()
    if not path:
        return False
    drive = os.path.splitdrive(os.path.abspath(path))[0]
    if drive and not os.path.isdir(drive + os.sep):
        return False
    try:
        os.makedirs(path, exist_ok=True)
        test = os.path.join(path, ".写入测试.tmp")
        with open(test, "w") as f:
            f.write("ok")
        os.remove(test)
        return True
    except OSError:
        return False

STATE_TEXT = {"wait": "○  等待", "start": "◐  下载中", "ok": "✔  完成", "empty": "–  无数据", "fail": "✘  失败",
              "stopped": "■  已停止", "cancel": "·  未执行", "skip": "↷  已存在·跳过"}
STATE_COLOR = {"wait": "#6B7280", "start": "#2563EB", "ok": "#15803D", "empty": "#A16207", "fail": "#DC2626",
               "stopped": "#EA580C", "cancel": "#9CA3AF", "skip": "#0F766E"}

# 配色
C = {
    "page": "#F2F4F8", "card": "#FFFFFF", "border": "#E3E8EF", "line": "#D5DCE6",
    "text": "#1F2937", "label": "#374151", "muted": "#7B8494", "faint": "#B6BECA",
    "accent": "#0F766E", "accent_hover": "#0B5F59", "accent_soft": "#E3F3F0", "accent_text": "#0F766E",
    "header": "#0E3B38", "header_sub": "#9CCFC6", "chip": "#1C5A54",
    "danger": "#DC2626", "danger_soft": "#FDECEC",
    "stripe": "#F8FAFC", "select": "#D5EFEA", "head": "#F6F8FB",
    "log_bg": "#151B23", "log_fg": "#C9D1D9", "log_time": "#6E7681",
}
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def _base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


SETTINGS_PATH = os.path.join(_base_dir(), "影子科技下载器设置.json")
LOG_PATH = os.path.join(_base_dir(), "影子科技下载器.log")


def _read_json(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            d = json.load(f)
            return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def load_settings():
    """读取设置；第一次运行时自动导入三个旧版下载器的设置（场列表、下载目录等）。"""
    s = _read_json(SETTINGS_PATH)
    if s:
        return s
    old_r = _read_json(os.path.join(_base_dir(), "下载器设置.json"))
    old_t = _read_json(os.path.join(_base_dir(), "猪只转移下载器设置.json"))
    old_s = _read_json(os.path.join(_base_dir(), "母猪档案下载器设置.json"))
    for o in (old_r, old_s, old_t):
        for k in ("headless", "keep_login"):
            if k in o and k not in s:
                s[k] = o[k]
    root = old_r.get("download_dir") or old_s.get("download_dir")
    if root:
        s["download_dir"] = root
    if old_r.get("fields"):
        s["report_fields"] = old_r["fields"]
    lists = {}
    if old_t.get("farms"):
        lists[LIST_TRANSFER] = old_t["farms"]
    if isinstance(old_s.get("stage_farm_lists"), dict):
        lists.update({k: v for k, v in old_s["stage_farm_lists"].items() if k in sw.STAGE_NAMES})
    if lists:
        s["lists"] = lists
    if old_t.get("last"):
        s["transfer_last"] = old_t["last"]
    if "split_month" in old_t:
        s["split_month"] = old_t["split_month"]
    if old_s.get("stages"):
        s["sow_stages"] = old_s["stages"]
    if old_s.get("stage_farms"):
        s["sow_stage_farms"] = old_s["stage_farms"]
    return s


def _valid_date(s):
    s = (s or "").strip()
    if not DATE_RE.match(s):
        return None
    try:
        return datetime.datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        return None


def month_last_day(y, m):
    return (datetime.date(y, m, 28) + datetime.timedelta(days=4)).replace(day=1) - datetime.timedelta(days=1)


def month_ranges(y1, m1, y2, m2, split):
    """起止年月 → [(开始日期, 结束日期)]；split=True 时每个月一段。"""
    start, end = datetime.date(y1, m1, 1), month_last_day(y2, m2)
    if not split:
        return [(start, end)]
    out, cur = [], start
    while cur <= end:
        last = month_last_day(cur.year, cur.month)
        out.append((cur, last))
        cur = last + datetime.timedelta(days=1)
    return out


def split_names(text):
    return [x.strip() for x in re.split(r"[;；,，、\n]+", text or "") if x.strip()]


def _open_path(path):
    try:
        if os.name == "nt":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        messagebox.showerror("打开失败", str(e))


def _result_info(info):
    """run_tasks 的结果 → (保存路径, 数量文字, 条数)。"""
    if isinstance(info, dict):
        path = info.get("path") or ""
        total, rows = info.get("total"), info.get("rows")
        n = total if total is not None else rows
        cnt = "" if n is None else f"{n:,}"
        if total is not None and rows is not None and rows != total:
            cnt += f"（文件 {rows:,}）"
        if info.get("missing"):                       # 系谱查询：有编码没查到
            cnt += f" · 未找到 {info['missing']} 个"
        elif (info.get("parts") or 1) > 1 and "found" not in info:
            cnt += f" · {info['parts']} 段"
        return path, cnt, n
    return str(info or ""), "", None


class TabBar(ttk.Frame):
    """扁平的下划线标签栏（替代 ttk.Notebook 的方框标签）。
    用法和 Notebook 一样：页面以本控件为父控件创建，再 add(page, text=...)；select(i) / index_current()。
    所有页面叠放在同一格里、只把选中的提到最上面，所以切换时高度不变、下面的内容不会跳动。"""

    def __init__(self, parent, fonts, on_change=None):
        super().__init__(parent)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        self._head = tk.Frame(self, bg=C["card"])
        self._head.grid(row=0, column=0, sticky="ew")
        tk.Frame(self, bg=C["border"], height=1).grid(row=1, column=0, sticky="ew")
        self._fonts, self._on_change = fonts, on_change
        self._tabs, self._pages, self._cur = [], [], None

    def add(self, page, text):
        i = len(self._pages)
        box = tk.Frame(self._head, bg=C["card"], cursor="hand2")
        box.pack(side="left", padx=(0, 6))
        lab = tk.Label(box, text=text, bg=C["card"], fg=C["muted"], font=self._fonts[0], padx=16, pady=7,
                       cursor="hand2")
        lab.pack()
        bar = tk.Frame(box, bg=C["card"], height=3)
        bar.pack(fill="x", padx=10)
        for w in (box, lab, bar):
            w.bind("<Button-1>", lambda e, i=i: self.select(i))
            w.bind("<Enter>", lambda e, i=i: self._hover(i, True))
            w.bind("<Leave>", lambda e, i=i: self._hover(i, False))
        self._tabs.append((lab, bar))
        self._pages.append(page)
        page.grid(row=2, column=0, sticky="nsew")
        if self._cur is None:
            self.select(0)
        else:
            self._pages[self._cur].tkraise()

    def _hover(self, i, on):
        if i != self._cur:
            self._tabs[i][0].configure(fg=C["accent_text"] if on else C["muted"])

    def select(self, i):
        i = int(i)
        if not 0 <= i < len(self._pages):
            return
        self._cur = i
        for k, (lab, bar) in enumerate(self._tabs):
            sel = k == i
            lab.configure(fg=C["accent_text"] if sel else C["muted"], font=self._fonts[1 if sel else 0])
            bar.configure(bg=C["accent"] if sel else C["card"])
        self._pages[i].tkraise()
        if self._on_change:
            self._on_change(i)

    def index_current(self):
        return self._cur or 0


class App:
    def __init__(self, root):
        self.root = root
        self.msg_q = queue.Queue()
        self.worker = None
        self.running = False
        self.stop_requested = False
        self.tasks = []
        self.run_list = []
        self.run_t0 = None
        self.cur_task = None
        self.cur_t0 = None
        self.done_count = 0
        self._last_tick = 0
        self.lock_widgets = []
        self.collapsed, self._card_apply = {}, {}
        self.auto_run = False                       # 本次运行是定时触发的（结束时不弹窗）
        self.exit_when_done = "--run-schedule" in sys.argv[1:]
        self._last_sched_check = 0
        self.scale = max(1.0, root.winfo_fpixels("1i") / 96.0)

        rotated = core.rotate_file(LOG_PATH)               # 日志超过 5 MB 自动换新文件，保留 3 份旧的
        self.log_file = open(LOG_PATH, "a", encoding="utf-8")
        cleaned = core.clean_debug()                       # 出错截图只留 30 天内、最多 200 个
        self.settings = load_settings()
        self.sched = sp.normalize(self.settings.get("schedule"))
        self.fav_dirs = [d for d in (self.settings.get("fav_dirs") or []) if isinstance(d, str) and d.strip()]
        saved = self.settings.get("lists") or {}
        self.lists = {LIST_REPORT: list(saved.get(LIST_REPORT) or [rp.DEFAULT_ORG]),
                      LIST_TRANSFER: list(saved.get(LIST_TRANSFER) or DEFAULT_TRANSFER_FARMS),
                      LIST_GENETIC: list(saved.get(LIST_GENETIC) or []),
                      LIST_PEDIGREE: list(saved.get(LIST_PEDIGREE) or [])}
        for st in sw.STAGE_NAMES:
            self.lists[st] = list(saved.get(st) or sw.CONFIG["FARMS"])

        root.title(APP_TITLE)
        self._init_style()
        self._build_ui()
        self._prefill_credentials()
        self._init_collapse()
        self._fit_window()
        self._bind_zoom()
        self.root.after(200, self._restore_sash)
        if rotated:
            self._append_log("日志文件超过 5 MB，旧日志已改名为「影子科技下载器.log.1」")
        if cleaned:
            self._append_log(f"已清理 debug 文件夹里 {cleaned} 个过期的出错截图 / 页面文件（保留 30 天内）")
        self._handle_cli()
        self.root.after(400, self._check_dir_on_start)
        self._pump()
        root.protocol("WM_DELETE_WINDOW", self._on_close)

    def px(self, n):
        return int(n * self.scale)

    # ================================================================ 样式 / 布局
    def _init_style(self):
        style = ttk.Style()
        style.theme_use("clam")                     # clam 可完全自定义颜色，Windows / 其他系统外观一致
        if sys.platform.startswith("win"):
            family = "Microsoft YaHei UI"
        else:
            family = tkfont.nametofont("TkDefaultFont").actual("family")
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            try:
                tkfont.nametofont(name).configure(family=family)
            except Exception:
                pass
        self.family = family
        self.zoom = min(max(float(self.settings.get("zoom", 1.0) or 1.0), 0.7), 1.6)
        self._fonts = {}

        def F(size=9, *w):
            """命名字体：缩放时只要改字号，所有用到它的控件一起变。"""
            key = (size,) + tuple(w)
            if key not in self._fonts:
                self._fonts[key] = tkfont.Font(family=family, size=max(6, round(size * self.zoom)),
                                               weight="bold" if "bold" in w else "normal")
            return self._fonts[key]
        self.F = F
        self._scale_text_fonts()
        self.root.configure(bg=C["page"])
        for k, v in {"*TCombobox*Listbox.background": C["card"], "*TCombobox*Listbox.foreground": C["text"],
                     "*TCombobox*Listbox.selectBackground": C["select"],
                     "*TCombobox*Listbox.selectForeground": C["text"], "*TCombobox*Listbox.font": F(9)}.items():
            self.root.option_add(k, v)

        style.configure(".", background=C["card"], foreground=C["text"], font=F(9), bordercolor=C["line"],
                        lightcolor=C["card"], darkcolor=C["card"], troughcolor=C["border"],
                        focuscolor=C["accent"], selectbackground=C["select"], selectforeground=C["text"],
                        insertcolor=C["text"])
        style.configure("TFrame", background=C["card"])
        style.configure("Page.TFrame", background=C["page"])
        style.configure("TLabel", background=C["card"], foreground=C["label"])
        style.configure("Page.TLabel", background=C["page"], foreground=C["label"])
        style.configure("CardTitle.TLabel", font=F(10, "bold"), foreground=C["text"])
        style.configure("Hint.TLabel", foreground=C["muted"], font=F(9))
        style.configure("Stat.TLabel", foreground=C["muted"], font=F(9))
        style.configure("StatStrong.TLabel", foreground=C["accent_text"], font=F(9, "bold"))
        style.configure("Title.TLabel", font=F(12, "bold"), foreground=C["text"])

        # 按钮
        style.configure("TButton", background=C["card"], foreground=C["label"], bordercolor=C["line"],
                        lightcolor=C["card"], darkcolor=C["card"], padding=(12, 4), font=F(9))
        style.map("TButton",
                  background=[("disabled", C["card"]), ("pressed", "#E9EEF4"), ("active", "#F2F5F9")],
                  foreground=[("disabled", C["faint"])],
                  bordercolor=[("disabled", C["border"]), ("focus", C["accent"]), ("active", "#B9C3D0")],
                  lightcolor=[("pressed", "#E9EEF4"), ("active", "#F2F5F9")],
                  darkcolor=[("pressed", "#E9EEF4"), ("active", "#F2F5F9")])
        style.configure("Accent.TButton", background=C["accent"], foreground="#FFFFFF", bordercolor=C["accent"],
                        lightcolor=C["accent"], darkcolor=C["accent"], padding=(14, 5), font=F(9, "bold"))
        style.map("Accent.TButton",
                  background=[("disabled", "#A9C9C5"), ("pressed", C["accent_hover"]), ("active", C["accent_hover"])],
                  foreground=[("disabled", "#F4F8F7")],
                  bordercolor=[("disabled", "#A9C9C5"), ("active", C["accent_hover"])],
                  lightcolor=[("disabled", "#A9C9C5"), ("active", C["accent_hover"])],
                  darkcolor=[("disabled", "#A9C9C5"), ("active", C["accent_hover"])])
        style.configure("Big.Accent.TButton", padding=(26, 8), font=F(10, "bold"))
        style.configure("Danger.TButton", foreground=C["danger"], bordercolor="#F3B4B4")
        style.map("Danger.TButton",
                  background=[("disabled", C["page"]), ("active", C["danger_soft"])],
                  foreground=[("disabled", C["faint"])],
                  bordercolor=[("disabled", C["border"]), ("active", C["danger"])])
        style.configure("Page.TButton", background=C["page"], lightcolor=C["page"], darkcolor=C["page"])
        style.map("Page.TButton", background=[("disabled", C["page"]), ("active", "#E7EBF1")])
        style.configure("Zoom.TButton", background=C["page"], lightcolor=C["page"], darkcolor=C["page"],
                        padding=(6, 3), width=-2)
        style.map("Zoom.TButton", background=[("active", "#E7EBF1")])
        style.configure("Chip.TButton", background="#F1F4F8", bordercolor="#F1F4F8", lightcolor="#F1F4F8",
                        darkcolor="#F1F4F8", foreground=C["label"], padding=(10, 3), font=F(9), width=-4)
        style.map("Chip.TButton",
                  background=[("disabled", "#F8FAFC"), ("active", C["accent_soft"])],
                  foreground=[("disabled", "#C3CAD4"), ("active", C["accent_text"])],
                  bordercolor=[("disabled", "#F8FAFC"), ("active", C["accent_soft"])],
                  lightcolor=[("disabled", "#F8FAFC"), ("active", C["accent_soft"])],
                  darkcolor=[("disabled", "#F8FAFC"), ("active", C["accent_soft"])])
        style.configure("ChipOn.TButton", background=C["accent_soft"], bordercolor=C["accent"],
                        lightcolor=C["accent_soft"], darkcolor=C["accent_soft"], foreground=C["accent_text"],
                        padding=(10, 3), font=F(9), width=-4)
        style.map("ChipOn.TButton", background=[("disabled", "#F8FAFC"), ("active", C["accent_soft"])],
                  foreground=[("disabled", "#C3CAD4")])

        # 输入框
        for w in ("TEntry", "TCombobox", "TSpinbox"):
            style.configure(w, fieldbackground=C["card"], background=C["card"], foreground=C["text"],
                            bordercolor=C["line"], lightcolor=C["card"], darkcolor=C["card"],
                            arrowcolor=C["muted"], padding=(6, 4))
            style.map(w, bordercolor=[("focus", C["accent"])], lightcolor=[("focus", C["accent"])],
                      fieldbackground=[("readonly", C["card"]), ("disabled", C["stripe"])],
                      foreground=[("disabled", C["faint"])])
        # 复选框 / 单选框：用抗锯齿图标代替 clam 主题自带的「X」
        size = min((16, 20, 24), key=lambda v: abs(v - 16 * self.scale))
        self._icons = {k[:-len(f"_{size}")]: tk.PhotoImage(data=v)
                       for k, v in ui_icons.ICONS.items() if k.endswith(f"_{size}")}
        for kind, w, el in (("cb", "TCheckbutton", "Checkbutton"), ("rb", "TRadiobutton", "Radiobutton")):
            ic = lambda st: self._icons[f"{kind}_{st}"]
            name = f"Flat.{el}.indicator"
            try:
                style.element_create(name, "image", ic("off"),
                                     ("disabled", "selected", ic("on_dis")), ("disabled", ic("off_dis")),
                                     ("selected", ic("on")), ("active", ic("hover")), sticky="w")
            except tk.TclError:
                pass                                   # 已创建过（重复打开窗口时）
            style.layout(w, [(f"{el}.padding", {"sticky": "nswe", "children": [
                (name, {"side": "left", "sticky": ""}),
                (f"{el}.focus", {"side": "left", "sticky": "w", "children": [
                    (f"{el}.label", {"sticky": "nswe"})]})]})])
            style.configure(w, background=C["card"], foreground=C["label"], padding=(0, 2), focuscolor=C["card"])
            style.map(w, background=[("active", C["card"])], foreground=[("disabled", C["faint"])])

        # 标签页
        style.configure("TNotebook", background=C["card"], bordercolor=C["border"], lightcolor=C["card"],
                        darkcolor=C["card"], tabmargins=(0, 0, 0, 0))
        style.configure("TNotebook.Tab", background="#F1F4F8", foreground=C["muted"], bordercolor=C["border"],
                        lightcolor="#F1F4F8", darkcolor="#F1F4F8", padding=(18, 7), font=F(9, "bold"))
        style.map("TNotebook.Tab",
                  background=[("selected", C["card"]), ("active", C["accent_soft"])],
                  foreground=[("selected", C["accent_text"]), ("active", C["accent_text"])],
                  lightcolor=[("selected", C["accent"])], bordercolor=[("selected", C["border"])])

        # 表格
        style.configure("Treeview", background=C["card"], fieldbackground=C["card"], foreground=C["text"],
                        bordercolor=C["border"], lightcolor=C["card"], darkcolor=C["card"],
                        rowheight=self._row_h(), font=F(9))
        style.map("Treeview", background=[("selected", C["select"])], foreground=[("selected", C["text"])])
        style.configure("Treeview.Heading", background=C["head"], foreground=C["muted"], font=F(9, "bold"),
                        bordercolor=C["border"], lightcolor=C["head"], darkcolor=C["head"], relief="flat",
                        padding=(8, 6))
        style.map("Treeview.Heading", background=[("active", "#EDF1F6")])

        # 滚动条 / 进度条 / 分隔条
        for name, trough, thumb in (("Vertical.TScrollbar", C["card"], "#CDD4DE"),
                                    ("Horizontal.TScrollbar", C["card"], "#CDD4DE"),
                                    ("Log.Vertical.TScrollbar", C["log_bg"], "#3A4350")):
            style.configure(name, background=thumb, troughcolor=trough, bordercolor=trough, lightcolor=thumb,
                            darkcolor=thumb, arrowcolor=C["muted"], gripcount=0, arrowsize=self.px(11),
                            relief="flat")
            style.map(name, background=[("active", "#9AA6B6")])
        style.configure("Accent.Horizontal.TProgressbar", troughcolor="#E1E6ED", background=C["accent"],
                        bordercolor="#E1E6ED", lightcolor=C["accent"], darkcolor=C["accent"],
                        thickness=self.px(8))
        style.configure("TPanedwindow", background=C["page"])
        style.configure("Sash", sashthickness=self.px(12), gripcount=0, background=C["page"],
                        lightcolor=C["page"], bordercolor=C["page"])

    def _card(self, parent, title, subtitle=None, collapse_key=None, summary_var=None):
        """白色卡片：标题行 + 内容区。返回 (卡片, 标题行, 内容区)。
        collapse_key：可折叠（点标题行折叠 / 展开，状态记在设置里）；summary_var：折叠时标题行显示的摘要。"""
        outer = tk.Frame(parent, bg=C["card"], highlightthickness=1, highlightbackground=C["border"],
                         highlightcolor=C["border"])
        head = ttk.Frame(outer)
        head.pack(fill="x", padx=16, pady=(9, 0))
        arrow = None
        if collapse_key:
            arrow = ttk.Label(head, text="▾", style="Hint.TLabel", width=2, cursor="hand2")
            arrow.pack(side="left")
        tk.Frame(head, bg=C["accent"], width=3, height=14).pack(side="left", padx=(0, 8))
        t = ttk.Label(head, text=title, style="CardTitle.TLabel", cursor="hand2" if collapse_key else "")
        t.pack(side="left")
        sub = None
        if subtitle or summary_var is not None:
            sub = ttk.Label(head, text=subtitle or "", style="Hint.TLabel")
            sub.pack(side="left", padx=(10, 0))
        body = ttk.Frame(outer, padding=(16, 6, 16, 11))
        body.pack(fill="both", expand=True)
        if collapse_key:
            pad = tk.Frame(outer, bg=C["card"], height=9)

            def apply(collapsed):
                self.collapsed[collapse_key] = collapsed
                arrow.configure(text="▸" if collapsed else "▾")
                if collapsed:
                    body.pack_forget()
                    pad.pack(fill="x")
                    if sub is not None and summary_var is not None:
                        sub.configure(textvariable=summary_var)
                else:
                    pad.pack_forget()
                    body.pack(fill="both", expand=True)
                    if sub is not None and summary_var is not None:
                        sub.configure(textvariable="")
                        sub.configure(text=subtitle or "")

            def toggle(_e=None):
                apply(not self.collapsed.get(collapse_key, False))
                self.root.after(50, self._restore_sash)
            for w in (arrow, t) + ((sub,) if sub is not None else ()):
                w.bind("<Button-1>", toggle)
            self._card_apply[collapse_key] = apply
        return outer, head, body

    def _init_collapse(self):
        """账号已填好时「账号与设置」默认折叠，给任务列表和日志腾出空间。"""
        saved = self.settings.get("collapsed") or {}
        have_cred = bool(self.user_var.get().strip() and self.pwd_var.get().strip())
        self._card_apply["settings"](saved.get("settings", have_cred))
        self._card_apply["add"](saved.get("add", False))
        for v in (self.user_var, self.dir_var, self.headless_var, self.keep_login_var, self.date_dirs_var):
            v.trace_add("write", lambda *a: self._update_settings_summary())
        self._update_settings_summary()

    def _update_settings_summary(self):
        flags = [n for n, v in (("后台运行", self.headless_var), ("记住登录", self.keep_login_var),
                                ("按年/月分文件夹", self.date_dirs_var)) if v.get()]
        self.settings_summary.set(f"账号 {self.user_var.get().strip() or '（未填）'}  ·  保存到 "
                                  f"{self.dir_var.get().strip() or default_root()}"
                                  + ("  ·  " + "、".join(flags) if flags else "") + "　（点击展开）")

    # ---------------------------------------------------------------- 缩放
    def _row_h(self):
        return max(18, int(self.px(28) * self.zoom))

    def _bind_zoom(self):
        for seq, d in (("<Control-plus>", 1), ("<Control-equal>", 1), ("<Control-KP_Add>", 1),
                       ("<Control-minus>", -1), ("<Control-KP_Subtract>", -1)):
            self.root.bind_all(seq, lambda e, d=d: self._zoom_step(d))
        self.root.bind_all("<Control-Key-0>", lambda e: self._set_zoom(1.0))
        self.root.bind_all("<Control-MouseWheel>", lambda e: self._zoom_step(1 if e.delta > 0 else -1))
        self.root.bind_all("<Control-Button-4>", lambda e: self._zoom_step(1))
        self.root.bind_all("<Control-Button-5>", lambda e: self._zoom_step(-1))
        self._update_zoom_label()

    ZOOMS = (0.7, 0.8, 0.9, 1.0, 1.1, 1.2, 1.3, 1.4, 1.5, 1.6)

    def _zoom_step(self, d):
        z = self.ZOOMS
        i = min(range(len(z)), key=lambda k: abs(z[k] - self.zoom))
        self._set_zoom(z[min(max(i + d, 0), len(z) - 1)])
        return "break"

    def _set_zoom(self, z):
        """界面缩放：字体、表格行高、列宽一起按比例变（Ctrl + 滚轮 / Ctrl + 加减号 / Ctrl + 0 还原）。"""
        if abs(z - self.zoom) < 1e-6:
            return
        self.zoom = z
        for key, f in self._fonts.items():
            f.configure(size=max(6, round(key[0] * z)))
        self._scale_text_fonts()
        ttk.Style().configure("Treeview", rowheight=self._row_h())
        for c, w in self.col_widths.items():
            self.tree.column(c, width=int(self.px(w) * z))
        self._fill_tree_width()
        self._update_zoom_label()
        self.root.after(80, self._restore_sash)

    TEXT_FONT_PT = 9     # 输入框 / 下拉框 / 数字框（TkTextFont）在 100% 时的字号

    def _scale_text_fonts(self):
        """ttk 的输入框、下拉框、数字框用的是系统命名字体 TkTextFont，不走自定义字体，要单独跟着缩放。"""
        for name in ("TkTextFont", "TkDefaultFont"):
            try:
                tkfont.nametofont(name).configure(size=max(6, round(self.TEXT_FONT_PT * self.zoom)))
            except tk.TclError:
                pass

    def _update_zoom_label(self):
        if hasattr(self, "zoom_var"):
            self.zoom_var.set(f"{round(self.zoom * 100)}%")

    def _build_header(self):
        bar = tk.Frame(self.root, bg=C["header"])
        bar.grid(row=0, column=0, sticky="ew")
        inner = tk.Frame(bar, bg=C["header"])
        inner.pack(fill="x", padx=18, pady=12)
        logo = tk.Canvas(inner, width=36, height=36, bg=C["header"], highlightthickness=0)
        logo.create_oval(2, 2, 34, 34, fill=C["accent"], outline="#2A9D90", width=2)
        logo.create_text(18, 18, text="猪", fill="#FFFFFF", font=(self.family, 13, "bold"))
        logo.pack(side="left", padx=(0, 12))
        txt = tk.Frame(inner, bg=C["header"])
        txt.pack(side="left")
        tk.Label(txt, text="养殖数据批量下载", bg=C["header"], fg="#FFFFFF",
                 font=self.F(14, "bold")).pack(anchor="w")
        tk.Label(txt, text="母猪综合查询 · 猪只转移 · 母猪档案 · 遗传性状 · 育种值 · 系谱　｜　一次登录，自动导出、整理、汇总",
                 bg=C["header"], fg=C["header_sub"], font=self.F(9)).pack(anchor="w")
        self.sched_var = tk.StringVar()
        self.sched_chip = tk.Label(inner, textvariable=self.sched_var, bg=C["chip"], fg="#E6F4F1",
                                   font=self.F(9), padx=12, pady=5, cursor="hand2")
        self.sched_chip.bind("<Button-1>", lambda e: self._edit_schedule())
        # 右上角：缩放和最适窗口（放在标题栏，给底部状态栏文字留出空间）
        zm = tk.Frame(inner, bg=C["header"])
        zm.pack(side="right", padx=(12, 0))

        def hbtn(text, cmd, padx=8):
            b = tk.Label(zm, text=text, bg=C["chip"], fg="#E6F4F1", font=self.F(9), padx=padx, pady=4,
                         cursor="hand2")
            b.bind("<Button-1>", lambda e: cmd())
            b.bind("<Enter>", lambda e: b.configure(bg=C["accent"]))
            b.bind("<Leave>", lambda e: b.configure(bg=C["chip"]))
            return b
        hbtn("－", lambda: self._zoom_step(-1)).pack(side="left")
        self.zoom_var = tk.StringVar()
        zl = tk.Label(zm, textvariable=self.zoom_var, bg=C["header"], fg="#E6F4F1", font=self.F(9), width=5,
                      cursor="hand2")
        zl.pack(side="left")
        zl.bind("<Button-1>", lambda e: self._set_zoom(1.0))
        hbtn("＋", lambda: self._zoom_step(1)).pack(side="left")
        hbtn("⤢ 最适窗口", self._best_fit, padx=10).pack(side="left", padx=(8, 0))

    def _build_ui(self):
        self.root.rowconfigure(1, weight=1)
        self.root.columnconfigure(0, weight=1)
        self._build_header()
        outer = ttk.Frame(self.root, style="Page.TFrame", padding=(16, 12, 16, 10))
        outer.grid(row=1, column=0, sticky="nsew")
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(2, weight=1)

        self._build_settings(outer)

        self.add_summary = tk.StringVar(value="已折叠，点这里展开继续添加任务")
        card, head, body = self._card(outer, "添加下载任务", "选好条件后点「＋ 添加到任务列表」，可以混合添加各种任务",
                                      collapse_key="add", summary_var=self.add_summary)
        card.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        body.columnconfigure(0, weight=1)
        self.nb = TabBar(body, (self.F(10), self.F(10, "bold")))
        self.nb.grid(row=0, column=0, sticky="ew")
        self._tab_report(self.nb)
        self._tab_transfer(self.nb)
        self._tab_sow(self.nb)
        self._tab_genetic(self.nb)
        self._tab_ebv(self.nb)
        self._tab_pedigree(self.nb)
        try:
            self.nb.select(int(self.settings.get("tab", 0)))
        except Exception:
            pass

        self.paned = ttk.PanedWindow(outer, orient="vertical")
        self.paned.grid(row=2, column=0, sticky="nsew")
        self._build_task_list(self.paned)
        self._build_log(self.paned)
        self._build_bottom_bar(outer)

    def _build_settings(self, outer):
        self.settings_summary = tk.StringVar()
        card, head, f = self._card(outer, "账号与设置", "所有功能共用；账号密码也可写在程序目录的 .env 中",
                                   collapse_key="settings", summary_var=self.settings_summary)
        card.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        f.columnconfigure(1, weight=1)
        f.columnconfigure(3, weight=1)
        self.user_var = tk.StringVar()
        self.pwd_var = tk.StringVar()
        self.dir_var = tk.StringVar(value=self.settings.get("download_dir") or default_root())
        self.headless_var = tk.BooleanVar(value=bool(self.settings.get("headless", False)))
        self.keep_login_var = tk.BooleanVar(value=bool(self.settings.get("keep_login", True)))
        self.date_dirs_var = tk.BooleanVar(value=bool(self.settings.get("date_folders", True)))

        ttk.Label(f, text="账号").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        e_user = ttk.Entry(f, textvariable=self.user_var)
        e_user.grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Label(f, text="密码").grid(row=0, column=2, sticky="w", padx=(20, 10), pady=4)
        self.pwd_entry = ttk.Entry(f, textvariable=self.pwd_var, show="•")
        self.pwd_entry.grid(row=0, column=3, sticky="ew", pady=4)
        self.pwd_btn = ttk.Button(f, text="显示", width=6, command=self._toggle_pwd)
        self.pwd_btn.grid(row=0, column=4, sticky="e", padx=(8, 0), pady=4)

        ttk.Label(f, text="保存到").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=4)
        row = ttk.Frame(f)
        row.grid(row=1, column=1, columnspan=4, sticky="ew", pady=4)
        row.columnconfigure(0, weight=1)
        # 可直接输入，也可从下拉里选常用目录
        e_dir = ttk.Combobox(row, textvariable=self.dir_var, height=15, postcommand=self._fill_dir_box)
        e_dir.grid(row=0, column=0, sticky="ew")
        e_dir.bind("<FocusOut>", self._save_dir_now)
        e_dir.bind("<Return>", self._save_dir_now)
        e_dir.bind("<<ComboboxSelected>>", lambda e: self._use_dir(self.dir_var.get()))
        self.dir_box = e_dir
        b_dir = ttk.Button(row, text="选择文件夹…", command=self._browse_dir)
        b_dir.grid(row=0, column=1, padx=(8, 0))
        self.fav_btn = ttk.Button(row, text="☆ 加入常用", width=10, command=self._toggle_fav)
        self.fav_btn.grid(row=0, column=2, padx=(8, 0))
        b_fav = ttk.Button(row, text="管理常用…", command=self._edit_favs)
        b_fav.grid(row=0, column=3, padx=(8, 0))
        ttk.Button(row, text="打开", command=self._open_dir).grid(row=0, column=4, padx=(8, 0))
        b_def = b_fav

        ttk.Label(f, text="常用目录").grid(row=3, column=0, sticky="nw", padx=(0, 10), pady=(10, 0))
        self.fav_row = ttk.Frame(f)
        self.fav_row.grid(row=3, column=1, columnspan=4, sticky="ew", pady=(6, 0))
        self.fav_chips = []
        self._fav_w = 0
        self.fav_row.bind("<Configure>", lambda e: (e.width != self._fav_w) and (
            setattr(self, "_fav_w", e.width) or self.root.after_idle(self._flow_fav_chips)))
        self.dir_var.trace_add("write", lambda *a: self._update_fav_btn())
        self._build_fav_chips()

        opts = ttk.Frame(f)
        opts.grid(row=2, column=1, columnspan=4, sticky="w", pady=(4, 0))
        c_hl = ttk.Checkbutton(opts, text="后台运行（不显示浏览器）", variable=self.headless_var)
        c_hl.pack(side="left")
        c_kl = ttk.Checkbutton(opts, text="记住登录状态", variable=self.keep_login_var)
        c_kl.pack(side="left", padx=(22, 0))
        c_dd = ttk.Checkbutton(opts, text="按年 / 月自动分文件夹", variable=self.date_dirs_var)
        c_dd.pack(side="left", padx=(22, 0))
        ttk.Label(opts, text="例：母猪综合查询\\配种\\2026年\\06月\\…", style="Hint.TLabel").pack(side="left", padx=(8, 0))
        self.lock_widgets += [e_user, self.pwd_entry, e_dir, b_dir, b_def, c_hl, c_kl, c_dd]

    # ---------------------------------------------------------------- 母猪综合查询
    def _tab_report(self, nb):
        f = ttk.Frame(nb, padding=(6, 14, 6, 2))
        nb.add(f, text=rp.TITLE)
        f.columnconfigure(1, weight=1)
        lp = {"sticky": "w", "padx": (0, 10), "pady": 3}

        ttk.Label(f, text="公司 / 场").grid(row=0, column=0, **lp)
        row = ttk.Frame(f)
        row.grid(row=0, column=1, sticky="ew")
        row.columnconfigure(0, weight=1)
        orgs = self.lists[LIST_REPORT]
        self.r_org = tk.StringVar(value=self.settings.get("report_org") or (orgs[0] if orgs else rp.DEFAULT_ORG))
        self.r_org_cb = ttk.Combobox(row, textvariable=self.r_org, height=25)
        self.r_org_cb.grid(row=0, column=0, sticky="ew")
        b = ttk.Button(row, text="编辑场列表…", command=lambda: self._edit_list(LIST_REPORT))
        b.grid(row=0, column=1, padx=(8, 0))
        self.r_info = ttk.Label(row, text="", style="Hint.TLabel", width=12)
        self.r_info.grid(row=0, column=2, padx=(8, 0))
        self.lock_widgets.append(b)

        ttk.Label(f, text="日期字段").grid(row=1, column=0, **lp)
        row = ttk.Frame(f)
        row.grid(row=1, column=1, sticky="w")
        saved = self.settings.get("report_fields") or ["配种日期"]
        self.rf_vars = {}
        for fld in rp.GUI_FIELDS:
            v = tk.BooleanVar(value=fld in saved)
            self.rf_vars[fld] = v
            ttk.Checkbutton(row, text=fld, variable=v).pack(side="left", padx=(0, 16))
        ttk.Label(row, text="可多选：同一日期范围，每个字段生成一个任务（平台每次只能查一个日期字段）",
                  style="Hint.TLabel").pack(side="left")

        ttk.Label(f, text="日期范围").grid(row=2, column=0, **lp)
        row = ttk.Frame(f)
        row.grid(row=2, column=1, sticky="w")
        last = self.settings.get("report_last") or {}
        self.r_start = tk.StringVar(value=last.get("start", ""))
        self.r_end = tk.StringVar(value=last.get("end", ""))
        e1 = ttk.Entry(row, textvariable=self.r_start, width=12, justify="center")
        e1.pack(side="left")
        ttk.Label(row, text="至").pack(side="left", padx=6)
        e2 = ttk.Entry(row, textvariable=self.r_end, width=12, justify="center")
        e2.pack(side="left")
        for e in (e1, e2):
            e.bind("<Return>", lambda ev: self._add_report())
        ttk.Separator(row, orient="vertical").pack(side="left", fill="y", padx=12, pady=2)
        for txt, fn in [("今天", self._r_today), ("近7天", self._r_7d), ("本月", self._r_month),
                        ("上月", self._r_last_month), ("本年至今", self._r_ytd), ("去年", self._r_last_year)]:
            ttk.Button(row, text=txt, style="Chip.TButton", command=fn).pack(side="left", padx=(0, 6))

        ttk.Label(f, text="按季度").grid(row=3, column=0, **lp)
        row = ttk.Frame(f)
        row.grid(row=3, column=1, sticky="w")
        self.r_year = tk.StringVar(value=str(datetime.date.today().year))
        sp = ttk.Spinbox(row, from_=2015, to=2100, width=6, textvariable=self.r_year, justify="center",
                         command=self._update_quarter_buttons)
        sp.pack(side="left")
        sp.bind("<KeyRelease>", lambda e: self._update_quarter_buttons())
        ttk.Label(row, text="年").pack(side="left", padx=(4, 10))
        self.q_btns = {}
        for q in (1, 2, 3, 4):
            b = ttk.Button(row, text=f"{'一二三四'[q - 1]}季度", style="Chip.TButton",
                           command=lambda q=q: self._r_quarter(q))
            b.pack(side="left", padx=2)
            self.q_btns[q] = b
        self.fy_btn = ttk.Button(row, text="全年", style="Chip.TButton", command=self._r_full_year)
        self.fy_btn.pack(side="left", padx=2)
        self._update_quarter_buttons()

        ttk.Label(f, text="已下载").grid(row=4, column=0, **lp)
        row = ttk.Frame(f)
        row.grid(row=4, column=1, sticky="w")
        ttk.Label(row, text="已下载过的整月自动跳过，但最近").pack(side="left")
        self.r_recent = tk.StringVar(value=str(self.settings.get("report_refresh_months", 1)))
        sp_r = ttk.Spinbox(row, from_=0, to=12, width=3, textvariable=self.r_recent, justify="center")
        sp_r.pack(side="left", padx=4)
        ttk.Label(row, text="个月仍重新下载（平台可能补录）").pack(side="left")
        self.r_force = tk.BooleanVar(value=False)
        c_f = ttk.Checkbutton(row, text="强制重新下载（不跳过）", variable=self.r_force)
        c_f.pack(side="left", padx=(20, 0))
        self.lock_widgets += [sp_r, c_f]

        row = ttk.Frame(f)
        row.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Label(row, text="含整月的范围（季度、去年、本年至今…）自动按月拆成多个任务；不足一个月的放在配种 / 分娩 / 断奶文件夹下；仍超 5 万条再分段",
                  style="Hint.TLabel").pack(side="left")
        b = ttk.Button(row, text="＋ 添加到任务列表", style="Accent.TButton", command=self._add_report)
        b.pack(side="right")
        self.lock_widgets.append(b)

    def _set_range(self, s, e):
        self.r_start.set(s.isoformat())
        self.r_end.set(e.isoformat())

    def _r_today(self):
        t = datetime.date.today()
        self._set_range(t, t)

    def _r_7d(self):
        t = datetime.date.today()
        self._set_range(t - datetime.timedelta(days=6), t)

    def _r_month(self):
        t = datetime.date.today()
        self._set_range(t.replace(day=1), t)

    def _r_last_month(self):
        end = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
        self._set_range(end.replace(day=1), end)

    def _r_ytd(self):
        t = datetime.date.today()
        self._set_range(datetime.date(t.year, 1, 1), t)

    def _r_last_year(self):
        y = datetime.date.today().year - 1
        self._set_range(datetime.date(y, 1, 1), datetime.date(y, 12, 31))

    def _year_value(self):
        try:
            y = int(self.r_year.get().strip())
            return y if 2000 <= y <= 2100 else None
        except (ValueError, tk.TclError):
            return None

    def _quarter_bounds(self, q):
        y = self._year_value()
        if y is None:
            return None, None
        return datetime.date(y, q * 3 - 2, 1), month_last_day(y, q * 3)

    def _update_quarter_buttons(self):
        today = datetime.date.today()
        for q, btn in self.q_btns.items():
            _, end = self._quarter_bounds(q)
            btn.state(["!disabled"] if end and today > end else ["disabled"])
        y = self._year_value()
        self.fy_btn.state(["!disabled"] if y and y <= today.year else ["disabled"])

    def _r_quarter(self, q):
        s, e = self._quarter_bounds(q)
        if s:
            self._set_range(s, e)

    def _r_full_year(self):
        y, t = self._year_value(), datetime.date.today()
        if y:
            self._set_range(datetime.date(y, 1, 1), t if y == t.year else datetime.date(y, 12, 31))

    def _add_report(self):
        if self.running:
            return
        fields = [f for f in rp.GUI_FIELDS if self.rf_vars[f].get()]
        if not fields:
            messagebox.showwarning("提示", "请至少勾选一个日期字段")
            return
        s, e = _valid_date(self.r_start.get()), _valid_date(self.r_end.get())
        if s is None or e is None:
            messagebox.showwarning("提示", "日期格式应为 YYYY-MM-DD，例如 2026-10-01")
            return
        if s > e:
            messagebox.showwarning("提示", "开始日期不能晚于结束日期")
            return
        orgs = self._parse_farms(self.r_org.get(), LIST_REPORT)
        if not orgs:
            messagebox.showwarning("提示", "请选择或输入公司 / 场（场列表为空时请先点「编辑场列表…」）")
            return
        today = datetime.date.today()
        if s > today:
            messagebox.showwarning("提示", "开始日期在今天之后，还没有数据")
            return
        if e > today:
            e = today                                   # 未来的日期没有数据：只下载到今天
        pieces = sp.month_pieces(s, e)                  # 含完整月就按月拆开，避免超过 5 万条
        total = len(orgs) * len(fields) * len(pieces)
        if total > 60 and not messagebox.askyesno("确认", f"将生成 {total} 个任务，确定添加吗？"):
            return
        force = bool(self.r_force.get())
        self._apply_refresh_months()
        added = dup = skipped = 0
        for org in orgs:
            for fld in fields:
                for a, b in pieces:
                    r = self._add_report_task(org, fld, a, b, force)
                    if r == "skip":
                        skipped += 1
                    elif r:
                        added += 1
                    else:
                        dup += 1
        self._remember_farms(LIST_REPORT, orgs)
        self._after_add(added, dup)
        extra = []
        if len(pieces) > 1:
            extra.append(f"{s} ~ {e} 已按月拆成 {len(pieces)} 段")
        if skipped:
            extra.append(f"{skipped} 个整月以前已下载，已跳过")
        if extra:
            self.status_var.set(self.status_var.get() + f"（{'；'.join(extra)}）")
        if force:
            self.r_force.set(False)                     # 强制只对这一次添加有效，避免下次忘了关

    def _apply_refresh_months(self):
        try:
            n = min(max(int(self.r_recent.get()), 0), 12)
        except (ValueError, tk.TclError):
            n = 1
        rp.CONFIG["REFRESH_RECENT_MONTHS"] = n
        tr.CONFIG["REFRESH_RECENT_MONTHS"] = n
        return n

    def _add_report_task(self, org, fld, a, b, force=False):
        """添加一个母猪综合查询任务。返回 True = 新添加；False = 列表里已有；"skip" = 以前已下载过，标为跳过。
        先按当前「保存到」目录检查是否已下载过（整月、且不是最近要刷新的月份才跳过）。"""
        tag = f"{a:%Y年%m月}" if sp.is_full_month(a, b) else f"{a} ~ {b}"
        t, new = self._add("report", f"{org} · {fld}  {tag}", org=org, field=fld,
                           start=a.isoformat(), end=b.isoformat())
        if not new and not (force and t["state"] == "skip"):
            return False
        info = None
        if not force:
            rp.CONFIG["DOWNLOAD_DIR"] = os.path.join(os.path.abspath(self.dir_var.get().strip() or default_root()),
                                                     rp.TITLE)
            rp.CONFIG["DATE_FOLDERS"] = bool(self.date_dirs_var.get())
            info = rp.skip_info(rp.make_group(fld, a.isoformat(), b.isoformat(), org))
        if info:
            path, n = info
            t.update(state="skip", path=path, n=None, count="" if n is None else f"{n:,}",
                     note="以前已下载；勾选「强制重新下载」可重下")
            self._render(t)
            return "skip"
        if t["state"] == "skip":                        # 强制：把之前标为跳过的任务改回等待
            t.update(state="wait", path="", count="", note="")
            self._render(t)
        return True

    # ---------------------------------------------------------------- 猪只转移
    def _tab_transfer(self, nb):
        today = datetime.date.today()
        last = self.settings.get("transfer_last") or {}
        f = ttk.Frame(nb, padding=(6, 14, 6, 2))
        nb.add(f, text=tr.TITLE)
        f.columnconfigure(1, weight=1)
        lp = {"sticky": "w", "padx": (0, 10), "pady": 3}

        ttk.Label(f, text="所在场").grid(row=0, column=0, **lp)
        row = ttk.Frame(f)
        row.grid(row=0, column=1, sticky="ew")
        row.columnconfigure(0, weight=1)
        lf = last.get("farm") if last.get("farm") not in (None, "", "当前场") else ""
        farms = self.lists[LIST_TRANSFER]
        self.t_farm = tk.StringVar(value=lf or (farms[0] if farms else ""))
        self.t_farm_cb = ttk.Combobox(row, textvariable=self.t_farm, height=25)
        self.t_farm_cb.grid(row=0, column=0, sticky="ew")
        b = ttk.Button(row, text="编辑场列表…", command=lambda: self._edit_list(LIST_TRANSFER))
        b.grid(row=0, column=1, padx=(8, 0))
        self.t_info = ttk.Label(row, text="", style="Hint.TLabel", width=12)
        self.t_info.grid(row=0, column=2, padx=(8, 0))
        self.lock_widgets.append(b)

        ttk.Label(f, text="转移年月").grid(row=1, column=0, **lp)
        row = ttk.Frame(f)
        row.grid(row=1, column=1, sticky="w")

        def ym(y, m):
            yv, mv = tk.StringVar(value=str(y)), tk.StringVar(value=f"{m}月")
            ttk.Spinbox(row, from_=2015, to=2100, width=6, textvariable=yv, justify="center").pack(side="left")
            ttk.Label(row, text="年").pack(side="left", padx=(4, 6))
            ttk.Combobox(row, textvariable=mv, values=MONTHS, state="readonly", width=6,
                         justify="center").pack(side="left")
            return yv, mv

        self.y1, self.m1 = ym(last.get("y1") or today.year, last.get("m1") or today.month)
        ttk.Label(row, text="至").pack(side="left", padx=12)
        self.y2, self.m2 = ym(last.get("y2") or today.year, last.get("m2") or today.month)
        self.range_lbl = ttk.Label(row, text="", style="Hint.TLabel")
        self.range_lbl.pack(side="left", padx=(16, 0))
        for v in (self.y1, self.m1, self.y2, self.m2):
            v.trace_add("write", lambda *a: self._show_range())

        ttk.Label(f, text="快捷").grid(row=2, column=0, **lp)
        row = ttk.Frame(f)
        row.grid(row=2, column=1, sticky="w")
        for txt, fn in [("本月", self._t_this_month), ("上月", self._t_last_month), ("近3个月", self._t_3m),
                        ("本年至今", self._t_ytd), ("去年全年", self._t_last_year)]:
            ttk.Button(row, text=txt, style="Chip.TButton", command=fn).pack(side="left", padx=(0, 6))
        ttk.Separator(row, orient="vertical").pack(side="left", fill="y", padx=8, pady=2)
        self.tq_btns = {}
        for q in (1, 2, 3, 4):
            b = ttk.Button(row, text=f"{'一二三四'[q - 1]}季度", style="Chip.TButton",
                           command=lambda q=q: self._t_quarter(q))
            b.pack(side="left", padx=(0, 4))
            self.tq_btns[q] = b
        ttk.Label(row, text="（按起始年份；未过完的季度不可选）", style="Hint.TLabel").pack(side="left", padx=(2, 0))

        row = ttk.Frame(f)
        row.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        self.split_var = tk.BooleanVar(value=bool(self.settings.get("split_month", True)))
        ttk.Checkbutton(row, text="按月拆分（每个月一个任务、一个文件夹；数据多时更稳）",
                        variable=self.split_var).pack(side="left")
        ttk.Label(row, text="  多个场用 ; 隔开", style="Hint.TLabel").pack(side="left")
        b = ttk.Button(row, text="＋ 添加到任务列表", style="Accent.TButton", command=self._add_transfer)
        row2 = ttk.Frame(f)
        row2.grid(row=4, column=0, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Label(row2, text="已下完的整月自动跳过，但最近").pack(side="left")
        sp_t = ttk.Spinbox(row2, from_=0, to=12, width=3, textvariable=self.r_recent, justify="center")
        sp_t.pack(side="left", padx=4)
        ttk.Label(row2, text="个月仍重新检查（与母猪综合查询共用）").pack(side="left")
        self.t_force = tk.BooleanVar(value=False)
        c_tf = ttk.Checkbutton(row2, text="强制重新检查（不跳过）", variable=self.t_force)
        c_tf.pack(side="left", padx=(20, 0))
        self.lock_widgets += [sp_t, c_tf]
        b.pack(side="right")
        self.lock_widgets.append(b)
        self._show_range()

    def _ym(self):
        try:
            y1, y2 = int(self.y1.get()), int(self.y2.get())
            m1, m2 = MONTHS.index(self.m1.get()) + 1, MONTHS.index(self.m2.get()) + 1
            if not (2000 <= y1 <= 2100 and 2000 <= y2 <= 2100):
                return None
            return y1, m1, y2, m2
        except (ValueError, tk.TclError):
            return None

    def _set_ym(self, y1, m1, y2, m2):
        self.y1.set(str(y1))
        self.m1.set(f"{m1}月")
        self.y2.set(str(y2))
        self.m2.set(f"{m2}月")

    def _update_t_quarters(self):
        """猪只转移的季度按钮：和母猪综合查询一样，还没到或还没过完的季度不能点。"""
        today = datetime.date.today()
        try:
            y = int(self.y1.get())
        except (ValueError, tk.TclError):
            y = None
        for q, b in getattr(self, "tq_btns", {}).items():
            ok = y is not None and 2000 <= y <= 2100 and today > month_last_day(y, q * 3)
            b.state(["!disabled"] if ok and not self.running else ["disabled"])

    def _show_range(self):
        self._update_t_quarters()
        ym = self._ym()
        if not ym:
            self.range_lbl.configure(text="年份格式不正确")
        elif (ym[0], ym[1]) > (ym[2], ym[3]):
            self.range_lbl.configure(text="起始年月不能晚于截止年月")
        else:
            self.range_lbl.configure(text=f"即 {datetime.date(ym[0], ym[1], 1)} 至 {month_last_day(ym[2], ym[3])}")

    def _t_this_month(self):
        t = datetime.date.today()
        self._set_ym(t.year, t.month, t.year, t.month)

    def _t_last_month(self):
        d = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
        self._set_ym(d.year, d.month, d.year, d.month)

    def _t_3m(self):
        t = datetime.date.today()
        y, m = t.year, t.month - 2
        if m < 1:
            y, m = y - 1, m + 12
        self._set_ym(y, m, t.year, t.month)

    def _t_ytd(self):
        t = datetime.date.today()
        self._set_ym(t.year, 1, t.year, t.month)

    def _t_last_year(self):
        y = datetime.date.today().year - 1
        self._set_ym(y, 1, y, 12)

    def _t_quarter(self, q):
        try:
            y = int(self.y1.get())
        except ValueError:
            return
        if datetime.date.today() <= month_last_day(y, q * 3):
            return
        self._set_ym(y, q * 3 - 2, y, q * 3)

    def _add_transfer(self):
        if self.running:
            return
        ym = self._ym()
        if not ym:
            messagebox.showwarning("提示", "请正确填写起始、截止年月")
            return
        y1, m1, y2, m2 = ym
        if (y1, m1) > (y2, m2):
            messagebox.showwarning("提示", "起始年月不能晚于截止年月")
            return
        farms = self._parse_farms(self.t_farm.get(), LIST_TRANSFER)
        if not farms:
            messagebox.showwarning("提示", "请选择或输入所在场（场列表为空时请先点「编辑场列表…」）")
            return
        ranges = month_ranges(y1, m1, y2, m2, bool(self.split_var.get()))
        total = len(farms) * len(ranges)
        if total > 60 and not messagebox.askyesno("确认", f"将生成 {total} 个任务，确定添加吗？"):
            return
        force = bool(self.t_force.get())
        self._apply_refresh_months()
        added = dup = skipped = 0
        for fm in farms:
            for a, b in ranges:
                r = self._add_transfer_task(fm, a, b, force)
                if r == "skip":
                    skipped += 1
                elif r:
                    added += 1
                else:
                    dup += 1
        self._remember_farms(LIST_TRANSFER, farms)
        self._after_add(added, dup)
        if skipped:
            self.status_var.set(self.status_var.get() + f"（{skipped} 个整月以前已下完，已跳过）")
        if force:
            self.t_force.set(False)

    def _add_transfer_task(self, fm, a, b, force=False):
        """添加一个猪只转移任务。返回 True / False（已在列表）/ "skip"（以前已完整下完的整月）。"""
        t, new = self._add("transfer", f"{fm}  {a} ~ {b}", farm=fm, start=a.isoformat(), end=b.isoformat())
        if not new and not (force and t["state"] == "skip"):
            return False
        info = None
        if not force:
            tr.CONFIG["DOWNLOAD_DIR"] = os.path.join(os.path.abspath(self.dir_var.get().strip() or default_root()),
                                                     tr.TITLE)
            core.CONFIG["DATE_FOLDERS"] = bool(self.date_dirs_var.get())
            info = tr.skip_info(tr.make_group(fm, a.isoformat(), b.isoformat()))
        if info:
            path, n = info
            t.update(state="skip", path=path, n=None, count="" if n is None else f"{n:,}",
                     note="以前已完整下完；勾选「强制重新检查」可重新搜索")
            self._render(t)
            return "skip"
        if t["state"] == "skip":
            t.update(state="wait", path="", count="", note="")
            self._render(t)
        return True

    # ---------------------------------------------------------------- 母猪档案
    def _tab_sow(self, nb):
        f = ttk.Frame(nb, padding=(6, 14, 6, 2))
        nb.add(f, text=sw.TITLE)
        f.columnconfigure(2, weight=1)
        saved_stages = self.settings.get("sow_stages") or sw.STAGE_NAMES
        saved_farms = self.settings.get("sow_stage_farms") or {}
        self.stage_vars, self.stage_farm_vars, self.stage_boxes, self.stage_infos = {}, {}, {}, {}
        for r, name in enumerate(sw.STAGE_NAMES):
            v = tk.BooleanVar(value=name in saved_stages)
            self.stage_vars[name] = v
            ttk.Checkbutton(f, text=name, variable=v, width=18).grid(row=r, column=0, sticky="w", pady=3)
            ttk.Label(f, text="所在场").grid(row=r, column=1, sticky="w", padx=(10, 8))
            fv = tk.StringVar(value=saved_farms.get(name, ALL_FARMS))
            self.stage_farm_vars[name] = fv
            box = ttk.Combobox(f, textvariable=fv, width=40)
            box.grid(row=r, column=2, sticky="ew", pady=3)
            self.stage_boxes[name] = box
            b = ttk.Button(f, text="编辑场列表…", command=lambda st=name: self._edit_list(st))
            b.grid(row=r, column=3, sticky="w", padx=(8, 0))
            info = ttk.Label(f, text="", style="Hint.TLabel", width=12)
            info.grid(row=r, column=4, sticky="w", padx=(8, 0))
            self.stage_infos[name] = info
            self.lock_widgets.append(b)
        n = len(sw.STAGE_NAMES)
        ttk.Label(f, text="可下拉选择或直接输入，多个场用 ; 隔开；选「（场列表中全部场）」= 该阶段场列表里的每个场各下载一次",
                  style="Hint.TLabel").grid(row=n, column=0, columnspan=5, sticky="w", pady=(2, 0))
        row = ttk.Frame(f)
        row.grid(row=n + 1, column=0, columnspan=5, sticky="ew", pady=(6, 0))
        ttk.Label(row, text="下载后自动在表格里追加「场」「阶段类型」列，并记录到「母猪档案下载记录.csv」",
                  style="Hint.TLabel").pack(side="left")
        b = ttk.Button(row, text="＋ 添加到任务列表", style="Accent.TButton", command=self._add_sow)
        b.pack(side="right")
        self.lock_widgets.append(b)
        self._refresh_farm_boxes()

    def _add_sow(self):
        if self.running:
            return
        stages = [s for s in sw.STAGE_NAMES if self.stage_vars[s].get()]
        if not stages:
            messagebox.showwarning("提示", "请至少勾选一个阶段")
            return
        plan = []
        for st in stages:
            farms = self._parse_farms(self.stage_farm_vars[st].get(), st)
            if not farms:
                messagebox.showwarning("提示", f"「{st}」没有可用的场：请直接输入场名，或点这一行的「编辑场列表…」添加")
                return
            plan += [(fm, st) for fm in farms]
            self._remember_farms(st, farms)
        added = sum(1 for fm, st in plan if self._add("sow", f"{fm} · {st}", farm=fm, stage=st)[1])
        self._after_add(added, len(plan) - added)

    # ---------------------------------------------------------------- 遗传性状查询
    def _tab_genetic(self, nb):
        last = self.settings.get("genetic_last") or {}
        f = ttk.Frame(nb, padding=(6, 14, 6, 2))
        nb.add(f, text=gn.TITLE)
        f.columnconfigure(1, weight=1)
        f.columnconfigure(3, weight=1)
        lp = {"sticky": "w", "padx": (0, 10), "pady": 3}

        ttk.Label(f, text="性状").grid(row=0, column=0, **lp)
        row = ttk.Frame(f)
        row.grid(row=0, column=1, sticky="w")
        traits = list(gn.TRAITS)
        self.g_trait = tk.StringVar(value=last.get("trait") if last.get("trait") in traits else traits[0])
        ttk.Combobox(row, textvariable=self.g_trait, values=traits, state="readonly", width=12).pack(side="left")
        ttk.Label(row, text="基因组选种 › 性状信息 › 遗传性状查询", style="Hint.TLabel").pack(side="left", padx=(10, 0))

        ttk.Label(f, text="测定猪场").grid(row=1, column=0, **lp)
        row = ttk.Frame(f)
        row.grid(row=1, column=1, sticky="ew")
        row.columnconfigure(0, weight=1)
        self.g_farm = tk.StringVar(value=last.get("farm") or G_ANY)
        self.g_farm_cb = ttk.Combobox(row, textvariable=self.g_farm, height=25)
        self.g_farm_cb.grid(row=0, column=0, sticky="ew")
        b = ttk.Button(row, text="编辑场列表…", command=lambda: self._edit_list(LIST_GENETIC))
        b.grid(row=0, column=1, padx=(8, 0))
        self.g_info = ttk.Label(row, text="", style="Hint.TLabel", width=10)
        self.g_info.grid(row=0, column=2, padx=(8, 0))
        self.lock_widgets.append(b)

        ttk.Label(f, text="结测日期").grid(row=2, column=0, **lp)
        row = ttk.Frame(f)
        row.grid(row=2, column=1, sticky="w")
        self.g_mode = tk.StringVar(value="range" if last.get("mode") == "range" else "none")
        ttk.Radiobutton(row, text="不限", value="none", variable=self.g_mode,
                        command=self._g_mode_changed).pack(side="left")
        ttk.Radiobutton(row, text="指定", value="range", variable=self.g_mode,
                        command=self._g_mode_changed).pack(side="left", padx=(12, 8))
        self.g_start = tk.StringVar(value=last.get("start", ""))
        self.g_end = tk.StringVar(value=last.get("end", ""))
        self.g_e1 = ttk.Entry(row, textvariable=self.g_start, width=12, justify="center")
        self.g_e1.pack(side="left")
        ttk.Label(row, text="至").pack(side="left", padx=6)
        self.g_e2 = ttk.Entry(row, textvariable=self.g_end, width=12, justify="center")
        self.g_e2.pack(side="left")
        self.g_split = tk.BooleanVar(value=bool(last.get("split", True)))
        ttk.Checkbutton(row, text="按月拆分", variable=self.g_split).pack(side="left", padx=(12, 0))

        ttk.Label(f, text="快捷").grid(row=3, column=0, **lp)
        row = ttk.Frame(f)
        row.grid(row=3, column=1, sticky="w")
        for txt, fn in [("本月", self._g_month), ("上月", self._g_last_month), ("近3个月", self._g_3m),
                        ("本年至今", self._g_ytd), ("去年", self._g_last_year)]:
            ttk.Button(row, text=txt, style="Chip.TButton", command=fn).pack(side="left", padx=(0, 6))

        ttk.Label(f, text="国标编码").grid(row=0, column=2, sticky="nw", padx=(24, 10), pady=3)
        box = tk.Frame(f, bg=C["card"], highlightthickness=1, highlightbackground=C["line"],
                       highlightcolor=C["accent"])
        box.grid(row=1, column=3, rowspan=3, sticky="nsew", pady=3)
        box.columnconfigure(0, weight=1)
        box.rowconfigure(0, weight=1)
        self.g_text = tk.Text(box, height=4, width=30, font=self.F(9), relief="flat", borderwidth=0,
                              highlightthickness=0, padx=6, pady=4, wrap="none", undo=True,
                              selectbackground=C["select"], background=C["card"], foreground=C["text"])
        self.g_text.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.g_text.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.g_text.configure(yscrollcommand=sb.set)
        self.g_text.bind("<<Modified>>", self._g_codes_changed)
        row = ttk.Frame(f)
        row.grid(row=0, column=3, sticky="ew")
        self.g_count = ttk.Label(row, text="", style="Hint.TLabel")
        self.g_count.pack(side="left")
        b1 = ttk.Button(row, text="清空", style="Chip.TButton", command=lambda: self.g_text.delete("1.0", "end"))
        b1.pack(side="right")
        b2 = ttk.Button(row, text="从文件导入…", style="Chip.TButton", command=self._g_import)
        b2.pack(side="right", padx=(0, 6))
        self.lock_widgets += [b1, b2]

        row = ttk.Frame(f)
        row.grid(row=4, column=0, columnspan=4, sticky="ew", pady=(6, 0))
        ttk.Label(row, text="三项都可留空（= 不限）；多个场用 ; 隔开；编码每行一个，超 400 个自动分批再合并；超 5 万条按结测日期拆分",
                  style="Hint.TLabel").pack(side="left")
        b = ttk.Button(row, text="＋ 添加到任务列表", style="Accent.TButton", command=self._add_genetic)
        b.pack(side="right")
        self.lock_widgets.append(b)
        self._g_mode_changed()
        self._g_codes_changed()
        self._refresh_farm_boxes()

    def _g_mode_changed(self):
        st = ["!disabled"] if self.g_mode.get() == "range" and not self.running else ["disabled"]
        for e in (self.g_e1, self.g_e2):
            e.state(st)

    def _g_set(self, s, e):
        self.g_mode.set("range")
        self._g_mode_changed()
        self.g_start.set(s.isoformat())
        self.g_end.set(e.isoformat())

    def _g_month(self):
        t = datetime.date.today()
        self._g_set(t.replace(day=1), t)

    def _g_last_month(self):
        end = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
        self._g_set(end.replace(day=1), end)

    def _g_3m(self):
        t = datetime.date.today()
        y, m = t.year, t.month - 2
        if m < 1:
            y, m = y - 1, m + 12
        self._g_set(datetime.date(y, m, 1), t)

    def _g_ytd(self):
        t = datetime.date.today()
        self._g_set(datetime.date(t.year, 1, 1), t)

    def _g_last_year(self):
        y = datetime.date.today().year - 1
        self._g_set(datetime.date(y, 1, 1), datetime.date(y, 12, 31))

    def _g_codes(self):
        return gn.parse_codes(self.g_text.get("1.0", "end"))

    def _g_codes_changed(self, _e=None):
        try:
            self.g_text.edit_modified(False)
        except tk.TclError:
            pass
        codes, bad = self._g_codes()
        n, size = len(codes), gn.CONFIG["CODE_BATCH"]
        if not n:
            txt = "不填 = 不限；可粘贴，每行一个"
        else:
            txt = f"共 {n} 个" + (f"，分 {-(-n // size)} 批导出后合并" if n > size else "")
        if bad:
            txt += f"；{len(bad)} 处不像编码，已忽略（如 {bad[0][:16]}）"
        self.g_count.configure(text=txt)

    def _g_import(self):
        if self.running:
            return
        p = filedialog.askopenfilename(parent=self.root, title="选择国标编码文件",
                                       filetypes=[("编码文件", "*.txt *.csv *.xlsx *.xls"), ("所有文件", "*.*")])
        if not p:
            return
        try:
            codes = gn.codes_from_file(p)
        except Exception as e:
            messagebox.showerror("导入失败", f"无法读取「{os.path.basename(p)}」：{e}")
            return
        if not codes:
            messagebox.showwarning("提示", "文件里没有找到国标编码（Excel 请把编码放在表头含「国标」的一列）")
            return
        old, _bad = self._g_codes()
        merged = old + [c for c in codes if c.upper() not in {x.upper() for x in old}]
        self.g_text.delete("1.0", "end")
        self.g_text.insert("1.0", "\n".join(merged))
        self._g_codes_changed()
        self.status_var.set(f"已从「{os.path.basename(p)}」导入 {len(codes)} 个国标编码（去重后共 {len(merged)} 个）")

    def _g_farms(self):
        text = self.g_farm.get().strip()
        if not text or text in (G_ANY, "不限", "全部场", gn.ALL_FARM_TEXT):
            return [""]
        return self._parse_farms(text, LIST_GENETIC)

    def _add_genetic(self):
        if self.running:
            return
        trait = self.g_trait.get()
        farms = self._g_farms()
        if not farms:
            messagebox.showwarning("提示", "测定猪场列表为空：请直接输入场名，或选「不限」")
            return
        codes, bad = self._g_codes()
        if self.g_mode.get() == "range":
            s, e = _valid_date(self.g_start.get()), _valid_date(self.g_end.get())
            if s is None or e is None:
                messagebox.showwarning("提示", "结测日期格式应为 YYYY-MM-DD，例如 2026-08-01；不限日期请选「不限」")
                return
            if s > e:
                messagebox.showwarning("提示", "开始日期不能晚于结束日期")
                return
            ranges = [(a.isoformat(), b.isoformat())
                      for a, b in (sp.split_months(s, e) if self.g_split.get() else [(s, e)])]
        else:
            ranges = [("", "")]
        if bad and not messagebox.askyesno("确认", f"国标编码里有 {len(bad)} 处不像编码（如「{bad[0][:20]}」），"
                                                     f"将忽略这些内容，只用 {len(codes)} 个编码。继续吗？"):
            return
        if farms == [""] and ranges == [("", "")] and not codes:
            if not messagebox.askyesno("确认", f"测定猪场、结测日期、国标编码都没有填：将导出全部{trait}数据。\n"
                                                "数据量很大（超过 5 万条会自动按结测日期分段，耗时较长），确定吗？"):
                return
        total = len(farms) * len(ranges)
        if total > 60 and not messagebox.askyesno("确认", f"将生成 {total} 个任务，确定添加吗？"):
            return
        added = dup = 0
        for fm in farms:
            for a, b in ranges:
                t = gn.make_task(fm, a, b, codes, trait)
                if self._add("genetic", gn.task_name(t), trait=trait, farm=fm, start=a, end=b, codes=tuple(codes))[1]:
                    added += 1
                else:
                    dup += 1
        self._remember_farms(LIST_GENETIC, [x for x in farms if x])
        self._after_add(added, dup)

    # ---------------------------------------------------------------- 育种值查询
    def _tab_ebv(self, nb):
        f = ttk.Frame(nb, padding=(6, 14, 6, 2))
        nb.add(f, text=ev.TITLE)
        f.columnconfigure(1, weight=1)
        f.rowconfigure(1, weight=1)
        lp = {"sticky": "nw", "padx": (0, 10), "pady": 3}
        ttk.Label(f, text="国标编码").grid(row=0, column=0, **lp)
        row = ttk.Frame(f)
        row.grid(row=0, column=1, sticky="ew")
        ttk.Label(row, text="基因组选种 › 育种值 › 育种值查询 › 按个体", style="Hint.TLabel").pack(side="left")
        self.e_count = ttk.Label(row, text="", style="Hint.TLabel")
        self.e_count.pack(side="left", padx=(16, 0))
        b1 = ttk.Button(row, text="清空", style="Chip.TButton", command=lambda: self.e_text.delete("1.0", "end"))
        b1.pack(side="right")
        b2 = ttk.Button(row, text="从文件导入…", style="Chip.TButton", command=self._e_import)
        b2.pack(side="right", padx=(0, 6))
        b3 = ttk.Button(row, text="复制「遗传性状」里的编码", style="Chip.TButton", command=self._e_copy_genetic)
        b3.pack(side="right", padx=(0, 6))
        self.lock_widgets += [b1, b2, b3]
        box = tk.Frame(f, bg=C["card"], highlightthickness=1, highlightbackground=C["line"],
                       highlightcolor=C["accent"])
        box.grid(row=1, column=1, sticky="nsew", pady=3)
        box.columnconfigure(0, weight=1)
        box.rowconfigure(0, weight=1)
        self.e_text = tk.Text(box, height=5, width=30, font=self.F(9), relief="flat", borderwidth=0,
                              highlightthickness=0, padx=6, pady=4, wrap="none", undo=True,
                              selectbackground=C["select"], background=C["card"], foreground=C["text"])
        self.e_text.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.e_text.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.e_text.configure(yscrollcommand=sb.set)
        self.e_text.bind("<<Modified>>", self._e_codes_changed)
        row = ttk.Frame(f)
        row.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Label(row, text="只按国标编码查询（必填），每行一个；超过 400 个自动分批导出再合并成一个文件；不涉及日期",
                  style="Hint.TLabel").pack(side="left")
        b = ttk.Button(row, text="＋ 添加到任务列表", style="Accent.TButton", command=self._add_ebv)
        b.pack(side="right")
        self.lock_widgets.append(b)
        self._e_codes_changed()

    def _e_codes(self):
        return gn.parse_codes(self.e_text.get("1.0", "end"))

    def _e_codes_changed(self, _e=None):
        try:
            self.e_text.edit_modified(False)
        except tk.TclError:
            pass
        codes, bad = self._e_codes()
        n, size = len(codes), ev.CONFIG["CODE_BATCH"]
        if not n:
            txt = "必填：粘贴国标编码，每行一个"
        else:
            txt = f"共 {n} 个" + (f"，分 {-(-n // size)} 批导出后合并" if n > size else "")
        if bad:
            txt += f"；{len(bad)} 处不像编码，已忽略（如 {bad[0][:16]}）"
        self.e_count.configure(text=txt)

    def _e_set_codes(self, codes, src):
        old, _bad = self._e_codes()
        seen = {x.upper() for x in old}
        merged = old + [c for c in codes if c.upper() not in seen]
        self.e_text.delete("1.0", "end")
        self.e_text.insert("1.0", "\n".join(merged))
        self._e_codes_changed()
        self.status_var.set(f"已从{src}导入 {len(codes)} 个国标编码（去重后共 {len(merged)} 个）")

    def _e_import(self):
        if self.running:
            return
        p = filedialog.askopenfilename(parent=self.root, title="选择国标编码文件",
                                       filetypes=[("编码文件", "*.txt *.csv *.xlsx *.xls"), ("所有文件", "*.*")])
        if not p:
            return
        try:
            codes = gn.codes_from_file(p)
        except Exception as e:
            messagebox.showerror("导入失败", f"无法读取「{os.path.basename(p)}」：{e}")
            return
        if not codes:
            messagebox.showwarning("提示", "文件里没有找到国标编码（Excel 请把编码放在表头含「国标」的一列）")
            return
        self._e_set_codes(codes, f"「{os.path.basename(p)}」")

    def _e_copy_genetic(self):
        if self.running:
            return
        codes, _bad = self._g_codes()
        if not codes:
            messagebox.showwarning("提示", "「遗传性状查询」页里还没有填国标编码")
            return
        self._e_set_codes(codes, "「遗传性状查询」页")

    def _add_ebv(self):
        if self.running:
            return
        codes, bad = self._e_codes()
        if not codes:
            messagebox.showwarning("提示", "育种值查询必须填国标编码：请粘贴（每行一个）或从文件导入")
            return
        if bad and not messagebox.askyesno("确认", f"国标编码里有 {len(bad)} 处不像编码（如「{bad[0][:20]}」），"
                                                     f"将忽略这些内容，只用 {len(codes)} 个编码。继续吗？"):
            return
        t = ev.make_task(codes)
        added = self._add("ebv", ev.task_name(t), codes=tuple(codes))[1]
        self._after_add(int(added), int(not added))

    # ---------------------------------------------------------------- 系谱查询
    def _tab_pedigree(self, nb):
        last = self.settings.get("pedigree_last") or {}
        f = ttk.Frame(nb, padding=(6, 14, 6, 2))
        nb.add(f, text=pd.TITLE)
        f.columnconfigure(1, weight=1)
        f.rowconfigure(1, weight=1)
        lp = {"sticky": "w", "padx": (0, 10), "pady": 3}

        ttk.Label(f, text="国标编码").grid(row=0, column=0, sticky="nw", padx=(0, 10), pady=3)
        row = ttk.Frame(f)
        row.grid(row=0, column=1, sticky="ew")
        ttk.Label(row, text="基因组选种 › 基础信息 › 系谱信息", style="Hint.TLabel").pack(side="left")
        self.p_count = ttk.Label(row, text="", style="Hint.TLabel")
        self.p_count.pack(side="left", padx=(16, 0))
        b1 = ttk.Button(row, text="清空", style="Chip.TButton", command=lambda: self.p_text.delete("1.0", "end"))
        b1.pack(side="right")
        b2 = ttk.Button(row, text="从文件导入…", style="Chip.TButton", command=self._p_import)
        b2.pack(side="right", padx=(0, 6))
        b3 = ttk.Button(row, text="复制「育种值」里的编码", style="Chip.TButton",
                        command=lambda: self._p_copy_from(self._e_codes, "育种值查询"))
        b3.pack(side="right", padx=(0, 6))
        self.lock_widgets += [b1, b2, b3]
        box = tk.Frame(f, bg=C["card"], highlightthickness=1, highlightbackground=C["line"],
                       highlightcolor=C["accent"])
        box.grid(row=1, column=1, sticky="nsew", pady=3)
        box.columnconfigure(0, weight=1)
        box.rowconfigure(0, weight=1)
        self.p_text = tk.Text(box, height=4, width=30, font=self.F(9), relief="flat", borderwidth=0,
                              highlightthickness=0, padx=6, pady=4, wrap="none", undo=True,
                              selectbackground=C["select"], background=C["card"], foreground=C["text"])
        self.p_text.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.p_text.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.p_text.configure(yscrollcommand=sb.set)
        self.p_text.bind("<<Modified>>", self._p_codes_changed)

        ttk.Label(f, text="场").grid(row=2, column=0, **lp)
        row = ttk.Frame(f)
        row.grid(row=2, column=1, sticky="ew")
        row.columnconfigure(0, weight=1)
        self.p_farm = tk.StringVar(value=last.get("farm") or ALL_FARMS)
        self.p_farm_cb = ttk.Combobox(row, textvariable=self.p_farm, height=25)
        self.p_farm_cb.grid(row=0, column=0, sticky="ew")
        b = ttk.Button(row, text="编辑场列表…", command=lambda: self._edit_list(LIST_PEDIGREE))
        b.grid(row=0, column=1, padx=(8, 0))
        self.p_info = ttk.Label(row, text="", style="Hint.TLabel", width=10)
        self.p_info.grid(row=0, column=2, padx=(8, 0))
        self.lock_widgets.append(b)

        ttk.Label(f, text="在场状态").grid(row=3, column=0, **lp)
        row = ttk.Frame(f)
        row.grid(row=3, column=1, sticky="w")
        st = last.get("status") if last.get("status") in pd.STATUS_ORDER else pd.STATUS_DEFAULT
        self.p_status = tk.StringVar(value=st)
        ttk.Combobox(row, textvariable=self.p_status, values=list(pd.STATUS_ORDER), state="readonly",
                     width=18).pack(side="left")
        ttk.Label(row, text="先按「在场」把场列表逐个查一遍，没找到的再按「离场」逐个查；查到的编码不再重复查",
                  style="Hint.TLabel").pack(side="left", padx=(10, 0))

        row = ttk.Frame(f)
        row.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Label(row, text="编码自动去重；每次最多 400 个，超出自动分批；每次查询的文件按「场_状态_批次」命名保留，最后合并成一个完整文件",
                  style="Hint.TLabel").pack(side="left")
        b = ttk.Button(row, text="＋ 添加到任务列表", style="Accent.TButton", command=self._add_pedigree)
        b.pack(side="right")
        self.lock_widgets.append(b)
        self._p_codes_changed()
        self._refresh_farm_boxes()

    def _p_codes(self):
        return gn.parse_codes(self.p_text.get("1.0", "end"))

    def _p_codes_changed(self, _e=None):
        try:
            self.p_text.edit_modified(False)
        except tk.TclError:
            pass
        raw = [x for x in re.split(r"[\s,，;；、|]+", self.p_text.get("1.0", "end")) if x.strip()]
        codes, bad = self._p_codes()
        dup = len(raw) - len(codes) - len(bad)
        if not codes:
            txt = "必填：粘贴国标编码，每行一个"
        else:
            txt = f"共 {len(codes)} 个" + (f"（已去掉重复 {dup} 个）" if dup > 0 else "") \
                + ("，每次 400 个分批查" if len(codes) > ev.CONFIG["CODE_BATCH"] else "")
        if bad:
            txt += f"；{len(bad)} 处不像编码，已忽略（如 {bad[0][:16]}）"
        self.p_count.configure(text=txt)

    def _p_set_codes(self, codes, src):
        old, _bad = self._p_codes()
        seen = {x.upper() for x in old}
        merged = old + [c for c in codes if c.upper() not in seen]
        self.p_text.delete("1.0", "end")
        self.p_text.insert("1.0", "\n".join(merged))
        self._p_codes_changed()
        self.status_var.set(f"已从{src}导入 {len(codes)} 个国标编码（去重后共 {len(merged)} 个）")

    def _p_import(self):
        if self.running:
            return
        p = filedialog.askopenfilename(parent=self.root, title="选择国标编码文件",
                                       filetypes=[("编码文件", "*.txt *.csv *.xlsx *.xls"), ("所有文件", "*.*")])
        if not p:
            return
        try:
            codes = gn.codes_from_file(p)
        except Exception as e:
            messagebox.showerror("导入失败", f"无法读取「{os.path.basename(p)}」：{e}")
            return
        if not codes:
            messagebox.showwarning("提示", "文件里没有找到国标编码（Excel 请把编码放在表头含「国标」的一列）")
            return
        self._p_set_codes(codes, f"「{os.path.basename(p)}」")

    def _p_copy_from(self, getter, name):
        if self.running:
            return
        codes, _bad = getter()
        if not codes:
            messagebox.showwarning("提示", f"「{name}」页里还没有填国标编码")
            return
        self._p_set_codes(codes, f"「{name}」页")

    def _add_pedigree(self):
        if self.running:
            return
        codes, bad = self._p_codes()
        if not codes:
            messagebox.showwarning("提示", "系谱查询必须填国标编码：请粘贴（每行一个）或从文件导入")
            return
        farms = self._parse_farms(self.p_farm.get(), LIST_PEDIGREE)
        if not farms:
            messagebox.showwarning("提示", "请选择或输入要查询的场（多个场用 ; 隔开），或先点「编辑场列表…」")
            return
        if bad and not messagebox.askyesno("确认", f"国标编码里有 {len(bad)} 处不像编码（如「{bad[0][:20]}」），"
                                                     f"将忽略这些内容，只用 {len(codes)} 个编码。继续吗？"):
            return
        status = self.p_status.get() if self.p_status.get() in pd.STATUS_ORDER else pd.STATUS_DEFAULT
        t = pd.make_task(codes, farms, status)
        added = self._add("pedigree", pd.task_name(t), codes=tuple(t["codes"]), farms=tuple(t["farms"]),
                          status=status)[1]
        self._remember_farms(LIST_PEDIGREE, farms)
        self._after_add(int(added), int(not added))

    # ---------------------------------------------------------------- 场列表
    def _choices(self, key):
        lst = self.lists[key]
        return ([ALL_FARMS] + lst) if len(lst) > 1 else list(lst)

    def _refresh_farm_boxes(self):
        self.r_org_cb.configure(values=self._choices(LIST_REPORT))
        self.r_info.configure(text=f"共 {len(self.lists[LIST_REPORT])} 个")
        self.t_farm_cb.configure(values=self._choices(LIST_TRANSFER))
        self.t_info.configure(text=f"共 {len(self.lists[LIST_TRANSFER])} 个场")
        for st in sw.STAGE_NAMES:
            self.stage_boxes[st].configure(values=[ALL_FARMS] + self.lists[st])
            n = len(self.lists[st])
            self.stage_infos[st].configure(text=f"共 {n} 个场" if n else "场列表为空")
        if hasattr(self, "g_farm_cb"):
            self.g_farm_cb.configure(values=[G_ANY] + self._choices(LIST_GENETIC))
            n = len(self.lists[LIST_GENETIC])
            self.g_info.configure(text=f"共 {n} 个场" if n else "")
        if hasattr(self, "p_farm_cb"):
            self.p_farm_cb.configure(values=self._choices(LIST_PEDIGREE))
            n = len(self.lists[LIST_PEDIGREE])
            self.p_info.configure(text=f"共 {n} 个场" if n else "场列表为空")

    def _parse_farms(self, text, key):
        text = (text or "").strip()
        lst = self.lists[key]
        if not text or text == ALL_FARMS:
            return list(lst) if text == ALL_FARMS or key not in (LIST_TRANSFER, LIST_REPORT) else []
        out = []
        for n in split_names(text):
            if n == ALL_FARMS:
                out += [x for x in lst if x not in out]
            elif n not in out and n != "当前场":
                out.append(n)
        return out

    def _remember_farms(self, key, farms):
        """新输入的场自动加入对应的场列表。"""
        new = [f for f in farms if f not in self.lists[key]]
        if new:
            self.lists[key] += new
            self._refresh_farm_boxes()
        self._save_settings()

    def _edit_list(self, key):
        if self.running:
            return
        dlg = tk.Toplevel(self.root)
        dlg.configure(bg=C["card"])
        dlg.title(f"编辑场列表 · {key}")
        dlg.transient(self.root)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding=16)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text=f"「{key}」的场列表", style="Title.TLabel").pack(anchor="w")
        ttk.Label(frm, text="每行一个场名，必须与页面上「测定猪场」下拉里的名称完全一致：" if key == LIST_GENETIC
                  else "每行一个场名，必须与页面右上角场选择框里的名称完全一致：").pack(anchor="w", pady=(4, 0))
        txt = tk.Text(frm, width=40, height=14, font=self.F(10), relief="flat", borderwidth=0,
                      highlightthickness=1, highlightbackground=C["line"], highlightcolor=C["accent"],
                      padx=8, pady=6, selectbackground=C["select"])
        txt.pack(fill="both", expand=True, pady=(6, 8))
        txt.insert("1.0", "\n".join(self.lists[key]))
        txt.focus_set()
        bar = ttk.Frame(frm)
        bar.pack(fill="x")

        def copy_from(src):
            txt.delete("1.0", "end")
            txt.insert("1.0", "\n".join(self.lists[src]))

        def save():
            names = []
            for n in split_names(txt.get("1.0", "end")):
                if n != ALL_FARMS and n not in names:
                    names.append(n)
            self.lists[key] = names
            self._refresh_farm_boxes()
            self._save_settings()
            self.status_var.set(f"「{key}」场列表已保存，共 {len(names)} 个场")
            dlg.destroy()

        for o in [k for k in self.lists if k != key]:
            ttk.Button(bar, text=f"复制「{o}」", command=lambda o=o: copy_from(o)).pack(side="left", padx=(0, 4))
        ttk.Button(bar, text="取消", command=dlg.destroy).pack(side="right")
        ttk.Button(bar, text="保存", style="Accent.TButton", command=save).pack(side="right", padx=(0, 8))
        dlg.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dlg.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dlg.winfo_height()) // 3
        dlg.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    # ================================================================ 任务列表 / 日志
    def _build_task_list(self, outer):
        card, head, f = self._card(outer, "下载任务", "双击已完成的任务可打开文件 / 文件夹")
        outer.add(card, weight=3)
        self.stat_var = tk.StringVar(value="暂无任务")
        ttk.Label(head, textvariable=self.stat_var, style="StatStrong.TLabel").pack(side="right")
        f.rowconfigure(0, weight=1)
        f.columnconfigure(0, weight=1)
        cols = ("seq", "kind", "desc", "status", "count", "elapsed", "file")
        heads = {"seq": "#", "kind": "功能", "desc": "任务内容", "status": "状态", "count": "数量（条）",
                 "elapsed": "耗时", "file": "保存位置 / 说明"}
        widths = {"seq": 42, "kind": 100, "desc": 320, "status": 90, "count": 120, "elapsed": 72, "file": 460}
        self.col_widths = widths
        box = tk.Frame(f, bg=C["card"], highlightthickness=1, highlightbackground=C["border"])
        box.grid(row=0, column=0, sticky="nsew")
        box.rowconfigure(0, weight=1)
        box.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(box, columns=cols, show="headings", height=4, selectmode="extended")
        for c in cols:
            self.tree.heading(c, text=heads[c], anchor="w" if c in ("file", "desc") else "center")
            # 列不自动挤压：内容放不下时出现左右滚动条（也可 Shift + 滚轮左右滑动）
            self.tree.column(c, width=int(self.px(widths[c]) * self.zoom), minwidth=self.px(36),
                             anchor="w" if c in ("file", "desc") else "center", stretch=False)
        for st, color in STATE_COLOR.items():
            self.tree.tag_configure(st, foreground=color)
        self.tree.tag_configure("odd", background=C["stripe"])
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Configure>", lambda e: self._fill_tree_width())
        self.tree.bind("<Shift-MouseWheel>", lambda e: self.tree.xview_scroll(-1 if e.delta > 0 else 1, "units"))
        self.tree.bind("<Shift-Button-4>", lambda e: self.tree.xview_scroll(-1, "units"))
        self.tree.bind("<Shift-Button-5>", lambda e: self.tree.xview_scroll(1, "units"))
        self.tree.bind("<Double-1>", self._on_tree_double)
        self.tree.bind("<Delete>", lambda e: self._del_selected())

        tb = ttk.Frame(f)
        tb.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        b1 = ttk.Button(tb, text="删除选中", command=self._del_selected)
        b1.pack(side="left")
        b2 = ttk.Button(tb, text="清除已完成", command=self._clear_done)
        b2.pack(side="left", padx=8)
        b3 = ttk.Button(tb, text="清空列表", command=self._clear_tasks)
        b3.pack(side="left")
        self.lock_widgets += [b1, b2, b3]
        self.count_var = tk.StringVar(value="")
        ttk.Label(tb, textvariable=self.count_var, style="Hint.TLabel").pack(side="right")

    def _fill_tree_width(self):
        """窗口比表格宽时，把最后一列拉宽填满；比表格窄时保持列宽、用左右滚动条。"""
        try:
            avail = self.tree.winfo_width()
            if avail <= 1:
                return
            others = sum(self.tree.column(c, "width") for c in self.col_widths if c != "file")
            base = int(self.px(self.col_widths["file"]) * self.zoom)
            self.tree.column("file", width=max(base, avail - others - 6))
        except tk.TclError:
            pass

    def _build_log(self, outer):
        card, head, f = self._card(outer, "运行日志")
        outer.add(card, weight=2)
        b = ttk.Button(head, text="清空", style="Chip.TButton", command=self._clear_log)
        b.pack(side="right")
        f.rowconfigure(0, weight=1)
        f.columnconfigure(0, weight=1)
        box = tk.Frame(f, bg=C["log_bg"])
        box.grid(row=0, column=0, sticky="nsew")
        box.rowconfigure(0, weight=1)
        box.columnconfigure(0, weight=1)
        self.log_text = tk.Text(box, height=4, state="disabled", wrap="word", relief="flat", borderwidth=0,
                                highlightthickness=0, font=self.F(9), background=C["log_bg"],
                                foreground=C["log_fg"], insertbackground=C["log_fg"], padx=12, pady=8,
                                selectbackground="#264F78", spacing1=2)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.log_text.yview, style="Log.Vertical.TScrollbar")
        sb.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=sb.set)
        self.log_text.tag_configure("time", foreground=C["log_time"])
        self.log_text.tag_configure("err", foreground="#FF7B72")
        self.log_text.tag_configure("ok", foreground="#56D364")
        self.log_text.tag_configure("warn", foreground="#E3B341")
        self.log_text.tag_configure("head", foreground="#79C0FF")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _build_bottom_bar(self, outer):
        bar = ttk.Frame(outer, style="Page.TFrame")
        bar.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        bar.columnconfigure(1, weight=1)
        prog = ttk.Frame(bar, style="Page.TFrame")
        prog.grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(prog, mode="determinate", length=self.px(220),
                                        style="Accent.Horizontal.TProgressbar")
        self.progress.pack(side="left")
        self.pct_var = tk.StringVar(value="")
        ttk.Label(prog, textvariable=self.pct_var, style="Page.TLabel", width=5).pack(side="left", padx=(8, 0))
        self.status_var = tk.StringVar(value="就绪")
        st = ttk.Label(bar, textvariable=self.status_var, style="Page.TLabel", justify="left")
        st.grid(row=0, column=1, sticky="ew", padx=8)
        # 文字太长时自动换成两行，不会被右边的按钮挡住
        st.bind("<Configure>", lambda e: st.configure(wraplength=max(e.width - 4, 80)))
        b_sch = ttk.Button(bar, text="⏰ 定时下载", style="Page.TButton", command=self._edit_schedule)
        b_sch.grid(row=0, column=2, padx=(0, 8))
        self.lock_widgets.append(b_sch)
        ttk.Button(bar, text="📂 打开下载目录", style="Page.TButton", command=self._open_dir).grid(
            row=0, column=3, padx=(0, 8))
        self.stop_btn = ttk.Button(bar, text="■ 停止", width=8, style="Danger.TButton", command=self._stop,
                                   state="disabled")
        self.stop_btn.grid(row=0, column=4, padx=(0, 8))
        self.start_btn = ttk.Button(bar, text="▶  开始下载", style="Big.Accent.TButton", command=self._start)
        self.start_btn.grid(row=0, column=5)

    def _fit_window(self):
        self.root.update_idletasks()
        sw_, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        w = min(max(self.px(1100), self.root.winfo_reqwidth()), sw_ - 40)
        h = max(min(self.px(1000), sh - 80), min(self.root.winfo_reqheight(), sh - 80))
        x, y = max((sw_ - w) // 2, 0), max((sh - h) // 2 - 20, 0)
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        self.root.minsize(min(self.px(900), sw_ - 40), min(self.px(660), sh - 80))

    @staticmethod
    def _work_area(root):
        """屏幕可用区域（去掉任务栏）：(x, y, 宽, 高)。"""
        if os.name == "nt":
            try:
                import ctypes
                from ctypes import wintypes
                r = wintypes.RECT()
                if ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(r), 0):   # SPI_GETWORKAREA
                    # 程序已声明 DPI 感知，得到的是物理像素，与 Tk 的坐标一致
                    return r.left, r.top, r.right - r.left, r.bottom - r.top
            except Exception:
                pass
        return 0, 0, root.winfo_screenwidth(), root.winfo_screenheight() - 60

    def _best_fit(self):
        """最适窗口：按屏幕可用区域调整窗口大小和位置 —— 高度占满（任务列表和日志能多显示几行），
        宽度按内容来；内容在当前缩放下放不下时自动把缩放调小一档，直到放得下。"""
        try:
            if self.root.state() == "zoomed":
                self.root.state("normal")
        except tk.TclError:
            pass
        try:
            self.root.attributes("-zoomed", False)      # Linux 下的最大化
        except tk.TclError:
            pass
        x0, y0, aw, ah = self._work_area(self.root)
        frame_w, frame_h = self.px(16), self.px(34)      # 窗口边框 + 标题栏
        max_w, max_h = aw - frame_w, ah - frame_h

        def need():
            """放得下所需的最小尺寸：任务列表 / 日志可以伸缩（各自有滚动条），只给它们留最小高度；
            宽度按其他区域（设置、添加任务、底部按钮栏）来算，表格太宽有横向滚动条。"""
            self.root.update_idletasks()
            outer = self.paned.master
            others = [w for w in outer.grid_slaves() if w is not self.paned]
            rw = max([w.winfo_reqwidth() for w in others] + [0]) + self.px(36)
            rh = self.root.winfo_reqheight() - self.paned.winfo_reqheight() + int(self.px(240) * self.zoom)
            return rw, rh

        # 1) 先回到 100%（之前因为屏幕小缩过的话）；2) 放不下就折叠「账号与设置」；3) 还放不下再逐档缩小
        target = max(z for z in self.ZOOMS if z <= max(self.zoom, 1.0))
        self._set_zoom(target)
        rw, rh = need()
        if (rh > max_h or rw > max_w) and not self.collapsed.get("settings"):
            self._card_apply["settings"](True)
            rw, rh = need()
        while (rw > max_w or rh > max_h) and self.zoom > self.ZOOMS[0]:
            self._zoom_step(-1)
            rw, rh = need()
        w = min(max(int(self.px(1280) * self.zoom), rw), max_w)
        h = max_h
        x = x0 + max((aw - w - frame_w) // 2, 0)
        self.root.geometry(f"{w}x{h}+{x}+{y0}")
        self.root.after(120, self._restore_sash)
        self.status_var.set("已调整为最适窗口" + (f"（缩放 {round(self.zoom * 100)}%）" if self.zoom < 1.0 else ""))

    def _row_values(self, t, seq):
        info = os.path.basename(t["path"]) if t["path"] else t.get("note", "")
        return (seq, KINDS[t["kind"]][0], t["desc"], STATE_TEXT.get(t["state"], t["state"]),
                t["count"], t["elapsed"], info)

    def _render(self, t):
        if t in self.tasks:
            i = self.tasks.index(t)
            self.tree.item(t["iid"], values=self._row_values(t, i + 1),
                           tags=(t["state"],) + (("odd",) if i % 2 else ()))

    def _refresh_count(self):
        n = len(self.tasks)
        if not n:
            self.stat_var.set("暂无任务")
            self.count_var.set("在上方选好条件，点「＋ 添加到任务列表」")
            return
        ok = sum(1 for t in self.tasks if t["state"] in ("ok", "empty"))
        skip = sum(1 for t in self.tasks if t["state"] == "skip")
        fail = sum(1 for t in self.tasks if t["state"] == "fail")
        rows = sum(t.get("n") or 0 for t in self.tasks if t["state"] == "ok")
        kinds = [KINDS[k][0] for k in KIND_ORDER if any(t["kind"] == k for t in self.tasks)]
        self.stat_var.set(f"共 {n} 个  ·  完成 {ok}" + (f"  ·  跳过 {skip}" if skip else "")
                          + (f"  ·  失败 {fail}" if fail else "")
                          + (f"  ·  已下载 {rows:,} 条" if rows else ""))
        self.count_var.set("包含：" + "、".join(kinds))

    def _add(self, kind, desc, **params):
        """返回 (任务, 是否新添加)；同样的任务已在列表里时返回已有的那个。"""
        key = (kind,) + tuple(sorted(params.items()))
        old = next((t for t in self.tasks if t["key"] == key), None)
        if old is not None:
            return old, False
        t = dict(params, kind=kind, key=key, desc=desc, state="wait", path="", count="", elapsed="", note="",
                 n=None)
        t["iid"] = self.tree.insert("", "end", values=self._row_values(t, len(self.tasks) + 1), tags=("wait",))
        self.tasks.append(t)
        return t, True

    def _after_add(self, added, dup):
        self._refresh_count()
        self._save_settings()
        self.status_var.set(f"已添加 {added} 个任务" + (f"，跳过 {dup} 个重复任务" if dup else ""))

    def _reseq(self):
        for t in self.tasks:
            self._render(t)
        self._refresh_count()

    def _del_selected(self):
        if self.running:
            return
        sel = set(self.tree.selection())
        if not sel:
            return
        self.tree.delete(*sel)
        self.tasks = [t for t in self.tasks if t["iid"] not in sel]
        self._reseq()

    def _clear_tasks(self):
        if self.running or not self.tasks:
            return
        self.tree.delete(*self.tree.get_children())
        self.tasks.clear()
        self._refresh_count()
        self.status_var.set("已清空任务列表")

    def _clear_done(self):
        if self.running:
            return
        done = [t for t in self.tasks if t["state"] in ("ok", "empty", "skip")]
        if done:
            self.tree.delete(*[t["iid"] for t in done])
            self.tasks = [t for t in self.tasks if t not in done]
            self._reseq()

    def _on_tree_double(self, _event):
        iid = self.tree.focus()
        t = next((t for t in self.tasks if t["iid"] == iid), None)
        if t and t["path"] and os.path.exists(t["path"]):
            _open_path(t["path"])

    # ================================================================ 设置
    def _sash_ratio(self):
        try:
            total = self.paned.winfo_height()
            return round(self.paned.sashpos(0) / total, 3) if total > 0 else self.settings.get("sash_ratio")
        except Exception:
            return self.settings.get("sash_ratio")

    def _restore_sash(self):
        """任务列表 / 日志的分隔位置：恢复上次拖动的比例，没有时任务列表占 58%。"""
        try:
            self.root.update_idletasks()
            total = self.paned.winfo_height()
            if total <= 0:
                return
            ratio = self.settings.get("sash_ratio")
            ratio = ratio if isinstance(ratio, (int, float)) and 0.15 < ratio < 0.9 else 0.58
            self.paned.sashpos(0, max(self.px(120), min(int(total * ratio), total - self.px(110))))
        except Exception:
            pass

    def _save_settings(self):
        ym = self._ym()
        try:
            tab = self.nb.index_current()
        except Exception:
            tab = 0
        self.settings.update({
            "download_dir": self.dir_var.get().strip(),
            "fav_dirs": list(self.fav_dirs),
            "report_refresh_months": self._apply_refresh_months() if hasattr(self, "r_recent") else 1,
            "headless": bool(self.headless_var.get()),
            "keep_login": bool(self.keep_login_var.get()),
            "date_folders": bool(self.date_dirs_var.get()),
            "sash_ratio": self._sash_ratio(),
            "zoom": self.zoom,
            "collapsed": dict(self.collapsed),
            "tab": tab,
            "lists": self.lists,
            "report_fields": [k for k, v in self.rf_vars.items() if v.get()],
            "report_org": self.r_org.get().strip(),
            "report_last": {"start": self.r_start.get().strip(), "end": self.r_end.get().strip()},
            "split_month": bool(self.split_var.get()),
            "transfer_last": {"farm": self.t_farm.get().strip(),
                              **({"y1": ym[0], "m1": ym[1], "y2": ym[2], "m2": ym[3]} if ym else {})},
            "sow_stages": [k for k, v in self.stage_vars.items() if v.get()],
            "sow_stage_farms": {k: v.get() for k, v in self.stage_farm_vars.items()},
            "pedigree_last": {"farm": self.p_farm.get().strip(), "status": self.p_status.get()}
            if hasattr(self, "p_farm") else {},
            "genetic_last": {"trait": self.g_trait.get(), "farm": self.g_farm.get().strip(),
                             "mode": self.g_mode.get(), "start": self.g_start.get().strip(),
                             "end": self.g_end.get().strip(), "split": bool(self.g_split.get())},
            "schedule": self.sched,
        })
        try:
            tmp = SETTINGS_PATH + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
            os.replace(tmp, SETTINGS_PATH)
        except Exception:
            pass

    def _prefill_credentials(self):
        try:
            u, p = core.load_credentials()
            self.user_var.set(u)
            self.pwd_var.set(p)
        except (SystemExit, Exception):
            pass

    def _toggle_pwd(self):
        hidden = bool(self.pwd_entry.cget("show"))
        self.pwd_entry.configure(show="" if hidden else "•")
        self.pwd_btn.configure(text="隐藏" if hidden else "显示")

    def _browse_dir(self):
        cur = self.dir_var.get().strip()
        start = cur if os.path.isdir(cur) else os.path.dirname(cur) if os.path.isdir(os.path.dirname(cur)) \
            else os.path.expanduser("~")
        d = filedialog.askdirectory(initialdir=start, title="选择默认保存目录", parent=self.root)
        if not d:
            return False
        d = os.path.normpath(d)
        if not dir_usable(d):
            messagebox.showwarning("提示", f"无法写入「{d}」，请换一个文件夹。")
            return False
        self.dir_var.set(d)
        self._save_dir_now()
        return True

    def _save_dir_now(self, *_):
        """保存目录一改就记住（下次打开程序就是它），不用等到点「开始下载」。"""
        d = self.dir_var.get().strip()
        if d and d != self.settings.get("download_dir"):
            self._save_settings()
            self.status_var.set(f"默认保存目录已设为：{d}")

    def _reset_dir(self):
        self.dir_var.set(default_root())
        self._save_dir_now()

    # ---------------------------------------------------------------- 常用下载目录
    FAV_MAX = 12

    @staticmethod
    def _same_dir(a, b):
        n = lambda p: os.path.normcase(os.path.normpath(os.path.abspath(p.strip()))) if p and p.strip() else ""
        return n(a) == n(b)

    def _is_fav(self, d):
        return any(self._same_dir(d, x) for x in self.fav_dirs)

    def _fill_dir_box(self):
        vals = list(self.fav_dirs)
        if not any(self._same_dir(default_root(), x) for x in vals):
            vals.append(default_root())
        self.dir_box.configure(values=vals)

    def _use_dir(self, d):
        """切换到某个常用目录：检查能否写入，然后立刻记住。"""
        if self.running:
            return
        d = os.path.normpath(d.strip())
        if not dir_usable(d):
            messagebox.showwarning("提示", f"「{d}」现在无法写入（可能没有这个盘、网络盘没连上或没有权限），"
                                         "请换一个目录。")
            self.dir_var.set(self.settings.get("download_dir") or default_root())
            return
        self.dir_var.set(d)
        self._save_dir_now()

    def _toggle_fav(self):
        if self.running:
            return
        d = self.dir_var.get().strip()
        if not d:
            return
        if self._is_fav(d):
            self.fav_dirs = [x for x in self.fav_dirs if not self._same_dir(d, x)]
            self.status_var.set(f"已从常用目录中移除：{d}")
        else:
            if not dir_usable(d):
                messagebox.showwarning("提示", f"「{d}」无法写入，不能加入常用目录。")
                return
            if len(self.fav_dirs) >= self.FAV_MAX:
                messagebox.showwarning("提示", f"常用目录最多 {self.FAV_MAX} 个，请先在「管理常用…」里删掉一些。")
                return
            self.fav_dirs.append(os.path.normpath(d))
            self.status_var.set(f"已加入常用目录：{d}")
        self._build_fav_chips()
        self._save_settings()

    def _fav_label(self, d):
        name = os.path.basename(os.path.normpath(d)) or d
        drive = os.path.splitdrive(os.path.abspath(d))[0]
        names = [os.path.basename(os.path.normpath(x)) for x in self.fav_dirs]
        if names.count(name) > 1:          # 同名文件夹：带上上一级目录区分
            name = os.path.join(os.path.basename(os.path.dirname(os.path.normpath(d))), name)
        return f"{drive} {name}".strip() if drive else name

    def _build_fav_chips(self):
        for w in self.fav_chips:
            w.destroy()
        self.fav_chips = []
        if not self.fav_dirs:
            w = ttk.Label(self.fav_row, text="还没有常用目录：选好「保存到」后点「☆ 加入常用」，以后一点就能切换",
                          style="Hint.TLabel")
            self.fav_chips.append(w)
        for d in self.fav_dirs:
            b = ttk.Button(self.fav_row, text=self._fav_label(d), style="Chip.TButton",
                           command=lambda d=d: self._use_dir(d))
            b.bind("<Enter>", lambda e, d=d: self.status_var.set(f"常用目录：{d}"))
            if self.running:
                b.state(["disabled"])
            self.fav_chips.append(b)
        self._update_fav_btn()
        self._flow_fav_chips()

    def _flow_fav_chips(self, _e=None):
        """常用目录按钮按可用宽度自动换行（目录多、名字长时不会超出窗口）。
        用 place 按坐标逐个摆放：grid 会让不同行的按钮按列对齐，反而把按钮挤出窗口。"""
        self.fav_row.update_idletasks()
        avail = self.fav_row.winfo_width()
        if avail <= 1:
            avail = self.root.winfo_width() - self.px(200)
        gap, vgap = self.px(6), self.px(4)
        x = y = row_h = 0
        for w in self.fav_chips:
            need_w, need_h = w.winfo_reqwidth(), w.winfo_reqheight()
            if x and x + need_w > avail:
                x, y, row_h = 0, y + row_h + vgap, 0
            w.place(x=x, y=y)
            x += need_w + gap
            row_h = max(row_h, need_h)
        h = y + row_h
        if h and int(self.fav_row.cget("height") or 0) != h:
            self.fav_row.configure(height=h)

    def _update_fav_btn(self):
        if not hasattr(self, "fav_btn"):
            return
        fav = self._is_fav(self.dir_var.get())
        self.fav_btn.configure(text="★ 已常用" if fav else "☆ 加入常用")
        cur = self.dir_var.get()
        for b, d in zip([w for w in self.fav_chips if isinstance(w, ttk.Button)], self.fav_dirs):
            b.configure(style="ChipOn.TButton" if self._same_dir(cur, d) else "Chip.TButton")

    def _edit_favs(self):
        if self.running:
            return
        dlg = tk.Toplevel(self.root)
        dlg.configure(bg=C["card"])
        dlg.title("管理常用下载目录")
        dlg.transient(self.root)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding=16)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="常用下载目录", style="Title.TLabel").pack(anchor="w")
        ttk.Label(frm, text=f"每行一个文件夹路径（最多 {self.FAV_MAX} 个），顺序就是按钮的顺序；可以是网络盘，如 \\\\服务器\\共享\\数据").pack(
            anchor="w", pady=(4, 0))
        txt = tk.Text(frm, width=70, height=10, font=self.F(10), relief="flat", borderwidth=0,
                      highlightthickness=1, highlightbackground=C["line"], highlightcolor=C["accent"],
                      padx=8, pady=6, selectbackground=C["select"], wrap="none")
        txt.pack(fill="both", expand=True, pady=(6, 8))
        txt.insert("1.0", "\n".join(self.fav_dirs))
        txt.focus_set()
        bar = ttk.Frame(frm)
        bar.pack(fill="x")

        def add_dir():
            d = filedialog.askdirectory(parent=dlg, title="选择要加入的常用目录",
                                        initialdir=self.dir_var.get().strip() or os.path.expanduser("~"))
            if d:
                cur = txt.get("1.0", "end").strip()
                txt.insert("end", ("\n" if cur else "") + os.path.normpath(d))

        def save():
            out, bad = [], []
            for line in txt.get("1.0", "end").splitlines():
                d = line.strip().strip('"')
                if not d or any(self._same_dir(d, x) for x in out):
                    continue
                d = os.path.normpath(d)
                (out if dir_usable(d) else bad).append(d)
            if bad and not messagebox.askyesno(
                    "确认", "下面这些目录现在无法写入（可能没有这个盘或网络盘没连上）：\n\n" + "\n".join(bad)
                    + "\n\n仍然保留它们吗？（选「否」则删掉）", parent=dlg):
                bad = []
            out = (out + bad)[:self.FAV_MAX]
            self.fav_dirs = out
            self._build_fav_chips()
            self._save_settings()
            self.status_var.set(f"常用目录已保存，共 {len(out)} 个")
            dlg.destroy()

        ttk.Button(bar, text="添加文件夹…", command=add_dir).pack(side="left")
        ttk.Button(bar, text="加入当前「保存到」", command=lambda: txt.insert(
            "end", ("\n" if txt.get("1.0", "end").strip() else "") + self.dir_var.get().strip())).pack(
            side="left", padx=(6, 0))
        ttk.Button(bar, text="取消", command=dlg.destroy).pack(side="right")
        ttk.Button(bar, text="保存", style="Accent.TButton", command=save).pack(side="right", padx=(0, 8))
        dlg.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dlg.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dlg.winfo_height()) // 3
        dlg.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _check_dir_on_start(self):
        """启动时检查保存目录：第一次使用时请用户确认 / 选择；目录所在的盘在这台电脑上不存在时提示重新选择。"""
        if self.exit_when_done or self.tasks:
            return                                         # 定时 / 命令行自动运行：不弹窗
        d = self.dir_var.get().strip()
        if not self.settings.get("download_dir"):
            if messagebox.askyesno("设置默认保存目录",
                                   f"下载的文件默认保存到：\n\n{d}\n\n是否要换一个位置？\n"
                                   "（以后也可以在「账号与设置 → 保存到」里随时修改）"):
                self._browse_dir()
            self.settings["download_dir"] = self.dir_var.get().strip()
            self._save_settings()
        elif not dir_usable(d):
            fb = default_root()
            if messagebox.askyesno("保存目录不可用",
                                   f"上次设置的保存目录「{d}」在这台电脑上不可用（可能没有这个盘）。\n\n"
                                   "现在重新选择一个吗？\n（选「否」则使用默认位置：" + fb + "）"):
                if self._browse_dir():
                    return
            self.dir_var.set(fb)
            self._save_dir_now()

    def _open_dir(self):
        d = self.dir_var.get().strip()
        if d and os.path.isdir(d):
            _open_path(d)
        else:
            messagebox.showinfo("提示", "下载目录还不存在，开始下载后会自动创建。")

    # ================================================================ 运行
    def _set_running(self, flag):
        self.running = flag
        for w in self.lock_widgets:
            try:
                w.state(["disabled"] if flag else ["!disabled"])
            except Exception:
                pass
        self.start_btn.state(["disabled"] if flag else ["!disabled"])
        self.stop_btn.state(["!disabled"] if flag else ["disabled"])
        self._update_quarter_buttons()
        self._update_t_quarters()
        self._g_mode_changed()
        for w in self.fav_chips:
            if isinstance(w, ttk.Button):
                w.state(["disabled"] if flag else ["!disabled"])
        try:
            self.g_text.configure(state="disabled" if flag else "normal")
            self.e_text.configure(state="disabled" if flag else "normal")
            self.p_text.configure(state="disabled" if flag else "normal")
        except tk.TclError:
            pass

    @staticmethod
    def _payload(t):
        if t["kind"] == "report":
            return rp.make_group(t["field"], t["start"], t["end"], t.get("org", ""))
        if t["kind"] == "transfer":
            return tr.make_group(t["farm"], t["start"], t["end"])
        if t["kind"] == "genetic":
            return gn.make_task(t["farm"], t["start"], t["end"], list(t["codes"]), t["trait"])
        if t["kind"] == "ebv":
            return ev.make_task(list(t["codes"]))
        if t["kind"] == "pedigree":
            return pd.make_task(list(t["codes"]), list(t["farms"]), t["status"])
        return {"farm": t["farm"], "stage": t["stage"]}

    def _start(self, only=None, auto=False):
        """only：只运行这些任务（定时下载用）；auto：定时触发，出问题只记日志、不弹窗。"""
        if self.running:
            return False

        def warn(msg):
            if auto:
                self._append_log("定时下载未能开始：" + msg)
            else:
                messagebox.showwarning("提示", msg)

        if only is not None:
            run = list(only)
        else:
            if not self.tasks:
                warn("请先在上方选好条件，点「添加到任务列表」")
                return False
            run = [t for t in self.tasks if t["state"] not in ("ok", "empty", "skip")]
            if not run:
                if not messagebox.askyesno("提示", "列表中的任务都已完成，是否全部重新运行？"):
                    return False
                run = list(self.tasks)
        if not run:
            warn("没有要运行的任务")
            return False
        user, pwd = self.user_var.get().strip(), self.pwd_var.get().strip()
        if not user or not pwd:
            warn("请填写账号和密码（或检查程序目录的 .env）")
            if not auto:
                self.pwd_entry.focus_set()
            return False
        root_dir = os.path.abspath(self.dir_var.get().strip() or default_root())
        if not dir_usable(root_dir):
            if auto:
                fb = default_root()
                if not dir_usable(fb):
                    warn(f"保存目录「{root_dir}」不可用，默认目录「{fb}」也不可用")
                    return False
                self._append_log(f"保存目录「{root_dir}」不可用（这台电脑可能没有该盘），本次改存到「{fb}」")
                root_dir = fb
            else:
                messagebox.showwarning("保存目录不可用",
                                       f"无法在「{root_dir}」保存文件（这台电脑可能没有这个盘，或没有写入权限）。\n"
                                       "请重新选择保存目录。")
                if not self._browse_dir():
                    return False
                root_dir = os.path.abspath(self.dir_var.get().strip())
        try:
            for kind in KIND_ORDER:
                os.makedirs(os.path.join(root_dir, KINDS[kind][0]), exist_ok=True)
        except OSError as e:
            warn(f"无法创建下载目录：{e}")
            return False
        self.auto_run = auto
        self.dir_var.set(root_dir)
        self._save_settings()

        # 同一功能的任务排在一起连续执行（减少来回切换页面），功能顺序按列表中首次出现的顺序
        order = list(dict.fromkeys(t["kind"] for t in run))
        run.sort(key=lambda t: order.index(t["kind"]))
        segments, offset = [], 0
        for kind in order:
            items = [t for t in run if t["kind"] == kind]
            segments.append((kind, [self._payload(t) for t in items], offset))
            offset += len(items)
            KINDS[kind][1].CONFIG["DOWNLOAD_DIR"] = os.path.join(root_dir, KINDS[kind][0])

        core.CONFIG["KEEP_LOGIN"] = bool(self.keep_login_var.get())
        core.CONFIG["DATE_FOLDERS"] = bool(self.date_dirs_var.get())
        core.HEADLESS = bool(self.headless_var.get())
        core.STOP_EVENT = threading.Event()
        core.PROGRESS_HOOK = lambda n, total, unit="条": self.msg_q.put(("progress", n, total, unit))
        tr.CONFIG["MANUAL_PARAMS"] = False

        for t in run:
            t.update(state="wait", path="", count="", elapsed="", note="", n=None)
            self._render(t)
        self.run_list = run
        self.done_count = 0
        self.progress.configure(maximum=len(run), value=0)
        self.pct_var.set("0%")
        self.run_t0 = time.time()
        self.cur_task = None
        self.stop_requested = False
        self._set_running(True)
        self._append_log(f"开始下载 {len(run)} 个任务（{'后台运行' if core.HEADLESS else '显示浏览器'}）")
        self.worker = threading.Thread(target=self._work, args=(user, pwd, segments), daemon=True)
        self.worker.start()
        return True

    def _work(self, user, pwd, segments):
        def hook(line):
            self.msg_q.put(("log", re.sub(r"^\[\d{2}:\d{2}:\d{2}\]\s*", "", line)))

        core.LOG_HOOK = hook
        n_ok = 0
        total = sum(len(p) for _, p, _ in segments)
        try:
            # 所有功能共用一个浏览器：一个功能做完只回首页，再进入下一个功能，不重启浏览器、不重新登录
            res = core.run_all([(KINDS[kind][1].ADAPTER, payloads) for kind, payloads, _ in segments], user, pwd,
                               status_fn=lambda i, n, st, info=None: self.msg_q.put(("status", i, st, info)))
            n_ok = len(res)
            self.msg_q.put(("done", n_ok, total))
        except core.StopRequested:
            self.msg_q.put(("done", n_ok, total))
        except core.LoginFatal as e:
            self.msg_q.put(("error", str(e), False))
        except SystemExit as e:
            self.msg_q.put(("error", str(e), False))
        except BaseException as e:
            self.msg_q.put(("error", f"{type(e).__name__}: {e}\n{traceback.format_exc()}", True))
        finally:
            core.LOG_HOOK = None
            core.PROGRESS_HOOK = None

    def _stop(self):
        if self.running and core.STOP_EVENT is not None:
            core.STOP_EVENT.set()
            self.stop_requested = True
            self.stop_btn.state(["disabled"])
            self.status_var.set("正在停止…")
            self._append_log("收到停止指令，当前步骤结束后退出（猪只转移已导出的单据下次会自动跳过）")
            self.root.after(10000, self._force_stop_if_needed)

    def _force_stop_if_needed(self):
        if self.running and self.worker and self.worker.is_alive():
            self._append_log("停止超时，强制关闭浏览器")
            threading.Thread(target=core.force_quit, daemon=True).start()

    def _finish_running(self):
        for t in self.run_list:
            if t["state"] in ("wait", "start"):
                t["state"] = "cancel" if t["state"] == "wait" else "stopped"
                self._render(t)
        self._set_running(False)
        self.cur_task = None
        core.STOP_EVENT = None
        self._refresh_count()

    def _tick(self):
        now = time.time()
        if not self.running or now - self._last_tick < 1:
            return
        self._last_tick = now
        if not self.stop_requested:
            total = len(self.run_list)
            if self.cur_task is None and self.done_count == 0:
                head = "正在启动浏览器并登录…"
            else:
                idx = self.run_list.index(self.cur_task) + 1 if self.cur_task in self.run_list else self.done_count
                head = f"正在下载第 {max(idx, 1)}/{total} 个任务"
                if self.cur_task is not None:
                    head += f"（{KINDS[self.cur_task['kind']][0]}）"
            self.status_var.set(f"{head} · 已用时 {core.fmt_secs(now - self.run_t0)}")
        if self.cur_task is not None and self.cur_t0:
            self.cur_task["elapsed"] = core.fmt_secs(now - self.cur_t0)
            self._render(self.cur_task)

    def _pump(self):
        try:
            while True:
                kind, *payload = self.msg_q.get_nowait()
                if kind == "log":
                    self._append_log(payload[0])
                elif kind == "status":
                    self._on_status(*payload)
                elif kind == "progress":
                    n, total, unit = payload
                    if self.cur_task is not None:
                        if unit == "段":
                            self.cur_task["note"] = f"超过 5 万条，分段下载中：已完成 {n}/{total} 段"
                        elif unit == "批":
                            self.cur_task["note"] = f"国标编码分批下载中：已完成 {n}/{total} 批"
                        elif unit == "轮":
                            self.cur_task["note"] = f"按场 / 在场状态循环查询：已完成 {n}/{total} 轮"
                        else:
                            self.cur_task["count"] = f"{n}/{total}" if total else str(n)
                            self.cur_task["note"] = f"进行中：第 {n}/{total or '?'} {unit}"
                        self._render(self.cur_task)
                elif kind == "done":
                    self._on_done(*payload)
                elif kind == "error":
                    text, detail = payload
                    self._finish_running()
                    self.status_var.set("出错")
                    self._append_log("发生错误：" + text)
                    if self.auto_run:
                        self._after_auto_run()
                    else:
                        messagebox.showerror("错误", text if not detail else text.splitlines()[0])
        except queue.Empty:
            pass
        self._tick()
        self._check_schedule()
        self.root.after(100, self._pump)

    def _on_status(self, i, state, info):
        if not 1 <= i <= len(self.run_list):
            return
        t = self.run_list[i - 1]
        if state == "start":
            self.cur_task, self.cur_t0 = t, time.time()
            t["note"] = "进行中…"
        else:
            if self.cur_t0 and self.cur_task is t:
                t["elapsed"] = core.fmt_secs(time.time() - self.cur_t0)
            if state == "ok":
                t["path"], cnt, t["n"] = _result_info(info)
                t["count"] = cnt or t["count"]
                t["note"] = ""
            elif state == "empty":
                t["count"], t["note"], t["n"] = "0", "无数据", 0
            elif state == "fail":
                t["note"] = (info or "")[:100]
            if state in ("ok", "empty", "fail"):
                self.done_count += 1
                self.progress.configure(value=self.done_count)
                self.pct_var.set(f"{self.done_count * 100 // max(len(self.run_list), 1)}%")
            if self.cur_task is t:
                self.cur_task = None
        t["state"] = state
        self._render(t)
        self._refresh_count()

    def _on_done(self, n_ok, total):
        stopped = self.stop_requested
        empty = sum(1 for t in self.run_list if t["state"] == "empty")
        fail = sum(1 for t in self.run_list if t["state"] == "fail")
        self._finish_running()
        dur = core.fmt_secs(time.time() - self.run_t0) if self.run_t0 else ""
        if stopped:
            self.status_var.set(f"已停止：完成 {n_ok + empty}/{total} 个任务")
            self._append_log("已停止下载")
            return
        summary = f"完成 {n_ok}/{total} 个任务" + (f"，无数据 {empty} 个" if empty else "") \
            + (f"，失败 {fail} 个" if fail else "")
        self.status_var.set(f"{summary}，用时 {dur}")
        self._append_log(f"全部结束：{summary}，用时 {dur}")
        if self.auto_run:
            self._after_auto_run()
            return
        msg = f"下载结束：{summary}，用时 {dur}。"
        if fail:
            msg += "\n再点「开始下载」会只重试未完成的任务。"
        if messagebox.askyesno("下载完成", msg + "\n\n是否打开下载目录？"):
            _open_path(self.dir_var.get().strip())

    # ================================================================ 定时下载
    def _after_auto_run(self):
        self.auto_run = False
        self._show_next_run()
        if self.exit_when_done:
            self._append_log("定时下载完成，程序自动退出")
            self.root.after(2000, self._on_close)

    def _show_next_run(self):
        nxt = sp.next_run(self.sched)
        if self.sched["mode"] == "off" or not self.sched["templates"] or not nxt:
            text = ""
        elif self.sched["mode"] == "windows":
            text = f"⏰  下次定时下载  {nxt:%m-%d %H:%M}  ·  Windows 计划任务"
        else:
            text = f"⏰  下次定时下载  {nxt:%m-%d %H:%M}  ·  需保持程序开着"
        if text != self.sched_var.get():
            self.sched_var.set(text)
            if text:
                self.sched_chip.pack(side="right")
            else:
                self.sched_chip.pack_forget()

    def _check_schedule(self):
        now = time.time()
        if now - self._last_sched_check < 15:
            return
        self._last_sched_check = now
        self._show_next_run()
        if not self.running and sp.is_due(self.sched):
            self._append_log(f"到达定时时间（{sp.describe(self.sched)}），开始定时下载")
            self._run_schedule(auto=True)

    def _run_schedule(self, auto=True):
        """按计划把任务模板展开成本次的任务（日期按规则现算），只运行这些任务。"""
        if not self.sched["templates"]:
            if not auto:
                messagebox.showwarning("提示", "定时计划里还没有任务模板")
            return False
        self.sched["last_run"] = datetime.datetime.now().isoformat(timespec="seconds")
        self._save_settings()
        a, b = sp.date_range(self.sched["rule"])
        self._append_log(f"定时下载：日期范围「{self.sched['rule']}」= {a} ~ {b}")
        objs = []
        self._apply_refresh_months()
        n_skip = 0
        for kind, params in sp.expand(self.sched["templates"], self.sched["rule"]):
            if kind == "report":
                a_, b_ = datetime.date.fromisoformat(params["start"]), datetime.date.fromisoformat(params["end"])
                if self._add_report_task(params["org"], params["field"], a_, b_) == "skip":
                    n_skip += 1
                    continue
                desc = f"{params['org']} · {params['field']}  " + \
                    (f"{a_:%Y年%m月}" if sp.is_full_month(a_, b_) else f"{a_} ~ {b_}")
            elif kind == "transfer":
                if self._add_transfer_task(params["farm"], datetime.date.fromisoformat(params["start"]),
                                           datetime.date.fromisoformat(params["end"])) == "skip":
                    n_skip += 1
                    continue
                desc = f"{params['farm']}  {params['start']} ~ {params['end']}"
            elif kind == "genetic":
                desc = gn.task_name(gn.make_task(params["farm"], params["start"], params["end"],
                                                 params["codes"], params["trait"]))
            elif kind == "ebv":
                desc = ev.task_name(ev.make_task(list(params["codes"])))
            elif kind == "pedigree":
                desc = pd.task_name(pd.make_task(list(params["codes"]), list(params["farms"]), params["status"]))
            else:
                desc = f"{params['farm']} · {params['stage']}"
            t, _ = self._add(kind, desc, **params)
            if t not in objs:
                objs.append(t)
        self._refresh_count()
        if n_skip:
            self._append_log(f"定时下载：{n_skip} 个整月以前已下载，跳过")
        if not objs:
            self._append_log("定时下载：没有需要下载的任务")
            return False
        return self._start(only=objs, auto=auto)

    def _templates_from_list(self):
        """把当前任务列表变成任务模板（去掉日期，同样的场 / 字段只留一个）。"""
        out = []
        for t in self.tasks:
            if t["kind"] == "report":
                tp = {"kind": "report", "org": t.get("org", ""), "field": t["field"]}
            elif t["kind"] == "transfer":
                tp = {"kind": "transfer", "farm": t["farm"], "split": bool(self.split_var.get())}
            elif t["kind"] == "genetic":
                tp = {"kind": "genetic", "trait": t["trait"], "farm": t["farm"], "codes": list(t["codes"]),
                      "dated": bool(t["start"]), "split": bool(self.g_split.get())}
            elif t["kind"] == "ebv":
                tp = {"kind": "ebv", "codes": list(t["codes"])}
            elif t["kind"] == "pedigree":
                tp = {"kind": "pedigree", "codes": list(t["codes"]), "farms": list(t["farms"]), "status": t["status"]}
            else:
                tp = {"kind": "sow", "farm": t["farm"], "stage": t["stage"]}
            if tp not in out:
                out.append(tp)
        return out

    def _edit_schedule(self):
        if self.running:
            return
        sc = dict(self.sched)
        sc["templates"] = list(sc["templates"])
        dlg = tk.Toplevel(self.root)
        dlg.configure(bg=C["card"])
        dlg.title("定时下载")
        dlg.transient(self.root)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding=16)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)
        lp = {"sticky": "w", "padx": (0, 10), "pady": 4}

        ttk.Label(frm, text="运行方式").grid(row=0, column=0, **lp)
        mode = tk.StringVar(value=sc["mode"])
        row = ttk.Frame(frm)
        row.grid(row=0, column=1, sticky="w")
        for k, txt in sp.MODES.items():
            ttk.Radiobutton(row, text=txt, value=k, variable=mode).pack(anchor="w")

        ttk.Label(frm, text="时间").grid(row=1, column=0, **lp)
        row = ttk.Frame(frm)
        row.grid(row=1, column=1, sticky="w")
        freq = tk.StringVar(value=sp.FREQS[sc["freq"]])
        wd = tk.StringVar(value=sp.WEEKDAYS[sc["weekday"]])
        md = tk.StringVar(value=str(sc["monthday"]))
        tm = tk.StringVar(value=sc["time"])
        ttk.Combobox(row, textvariable=freq, values=list(sp.FREQS.values()), state="readonly", width=6).pack(side="left")
        cb_wd = ttk.Combobox(row, textvariable=wd, values=sp.WEEKDAYS, state="readonly", width=5)
        sp_md = ttk.Spinbox(row, from_=1, to=28, width=4, textvariable=md, justify="center")
        lb_md = ttk.Label(row, text="日")
        lb_tm = ttk.Label(row, text="时刻")
        e_tm = ttk.Entry(row, textvariable=tm, width=7, justify="center")

        def relayout(*_):
            for w in (cb_wd, sp_md, lb_md, lb_tm, e_tm):
                w.pack_forget()
            f = next(k for k, v in sp.FREQS.items() if v == freq.get())
            if f == "weekly":
                cb_wd.pack(side="left", padx=(6, 0))
            elif f == "monthly":
                sp_md.pack(side="left", padx=(6, 0))
                lb_md.pack(side="left", padx=(2, 0))
            lb_tm.pack(side="left", padx=(14, 4))
            e_tm.pack(side="left")
        freq.trace_add("write", relayout)
        relayout()
        ttk.Label(frm, text="时刻格式 HH:MM，如 07:30；每月只能选 1–28 日", style="Hint.TLabel").grid(
            row=2, column=1, sticky="w")

        ttk.Label(frm, text="日期范围").grid(row=3, column=0, **lp)
        rule = tk.StringVar(value=sc["rule"])
        row = ttk.Frame(frm)
        row.grid(row=3, column=1, sticky="w")
        ttk.Combobox(row, textvariable=rule, values=sp.RULES, state="readonly", width=18).pack(side="left")
        rng_lbl = ttk.Label(row, text="", style="Hint.TLabel")
        rng_lbl.pack(side="left", padx=(10, 0))

        def show_rng(*_):
            a, b = sp.date_range(rule.get())
            rng_lbl.configure(text=f"今天运行的话 = {a} ~ {b}（母猪档案、育种值、不限日期的遗传性状不按日期）")
        rule.trace_add("write", show_rng)
        show_rng()

        ttk.Label(frm, text="任务模板").grid(row=4, column=0, sticky="nw", padx=(0, 10), pady=4)
        box = ttk.Frame(frm)
        box.grid(row=4, column=1, sticky="nsew", pady=4)
        box.columnconfigure(0, weight=1)
        frm.rowconfigure(4, weight=1)
        lst = tk.Listbox(box, height=8, font=self.F(9), activestyle="none", selectmode="extended",
                         relief="flat", borderwidth=0, highlightthickness=1, highlightbackground=C["line"],
                         highlightcolor=C["accent"], selectbackground=C["select"], selectforeground=C["text"],
                         background=C["card"], foreground=C["text"])
        lst.grid(row=0, column=0, sticky="nsew")

        def fill():
            lst.delete(0, "end")
            for t in sc["templates"]:
                lst.insert("end", sp.template_desc(t))
        fill()
        bar = ttk.Frame(box)
        bar.grid(row=1, column=0, sticky="ew", pady=(4, 0))

        def from_list():
            tps = self._templates_from_list()
            if not tps:
                messagebox.showwarning("提示", "任务列表是空的：先在主界面添加任务（日期随便选），再点这里", parent=dlg)
                return
            sc["templates"] = tps
            fill()

        def remove():
            for i in sorted(lst.curselection(), reverse=True):
                del sc["templates"][i]
            fill()
        ttk.Button(bar, text="用当前任务列表生成", command=from_list).pack(side="left")
        ttk.Button(bar, text="删除选中", command=remove).pack(side="left", padx=6)
        ttk.Label(frm, text="任务模板只记录「功能 + 公司/场 + 字段/阶段」，不记日期；\n每次运行时按上面的日期范围现算。"
                           "先在主界面把要定时下载的任务加到列表（日期随便选），再点「用当前任务列表生成」。",
                  style="Hint.TLabel").grid(row=5, column=1, sticky="w")

        def collect():
            t = sp.parse_time(tm.get())
            if t is None:
                messagebox.showwarning("提示", "时刻格式应为 HH:MM，例如 07:30", parent=dlg)
                return None
            sc.update(mode=mode.get(), freq=next(k for k, v in sp.FREQS.items() if v == freq.get()),
                      weekday=sp.WEEKDAYS.index(wd.get()), time=t.strftime("%H:%M"), rule=rule.get())
            try:
                sc["monthday"] = min(max(int(md.get()), 1), 28)
            except ValueError:
                sc["monthday"] = 1
            if sc["mode"] != "off" and not sc["templates"]:
                messagebox.showwarning("提示", "请先点「用当前任务列表生成」添加任务模板", parent=dlg)
                return None
            return sp.normalize(sc)

        def save():
            new = collect()
            if new is None:
                return
            if new["mode"] == "windows":
                ok, msg = sp.register_windows_task(new, __file__)
                if not ok:
                    messagebox.showerror("设置失败", f"未能写入 Windows 计划任务：\n{msg}", parent=dlg)
                    return
                self._append_log(f"已写入 Windows 计划任务「{sp.TASK_NAME}」：{sp.describe(new)}")
            elif sp.windows_task_exists():
                ok, msg = sp.delete_windows_task()
                self._append_log("已删除 Windows 计划任务" if ok else f"删除 Windows 计划任务失败：{msg}")
            if new["mode"] == "app" and sp.is_due(new):
                # 刚设置的时间今天已经过了：不要立刻补跑，从下一次开始
                new["last_run"] = datetime.datetime.now().isoformat(timespec="seconds")
            self.sched = new
            self._save_settings()
            self._show_next_run()
            self.status_var.set("定时下载已保存：" + sp.describe(new))
            dlg.destroy()

        def run_now():
            new = collect()
            if new is None:
                return
            if not new["templates"]:
                messagebox.showwarning("提示", "请先添加任务模板", parent=dlg)
                return
            self.sched = new
            self._save_settings()
            dlg.destroy()
            self._run_schedule(auto=False)

        btns = ttk.Frame(frm)
        btns.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(btns, text="按此计划立即运行一次", command=run_now).pack(side="left")
        ttk.Button(btns, text="取消", command=dlg.destroy).pack(side="right")
        ttk.Button(btns, text="保存", style="Accent.TButton", command=save).pack(side="right", padx=(0, 8))
        dlg.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dlg.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dlg.winfo_height()) // 3
        dlg.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _append_log(self, line):
        full = f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {line}"
        tag = ()
        if re.search(r"失败|错误|异常|超时", line):
            tag = ("err",)
        elif re.search(r"警告|无数据", line):
            tag = ("warn",)
        elif re.search(r"\[完成\]|成功|已保存|已生成", line):
            tag = ("ok",)
        if not tag and (line.startswith("====") or line.startswith("----") or line.startswith("== ")):
            tag = ("head",)
        self.log_text.configure(state="normal")
        self.log_text.insert("end", full[:10] + " ", ("time",))
        self.log_text.insert("end", line + "\n", tag)
        if int(self.log_text.index("end-1c").split(".")[0]) > 5000:
            self.log_text.delete("1.0", "1000.0")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        try:
            self.log_file.write(full + "\n")
            self.log_file.flush()
        except Exception:
            pass

    def _on_close(self):
        if self.running:
            if not messagebox.askyesno("退出", "下载正在进行，确定要退出吗？\n（退出后本次下载将中断）"):
                return
            if core.STOP_EVENT is not None:
                core.STOP_EVENT.set()
            # 非守护线程：确保看不见的后台浏览器被关掉后进程才退出
            threading.Thread(target=core.force_quit, daemon=False).start()
        self._save_settings()
        try:
            self.log_file.close()
        except Exception:
            pass
        self.root.destroy()

    # ================================================================ 命令行自动模式
    def _handle_cli(self):
        args = sys.argv[1:]
        if "--headless" in args:
            self.headless_var.set(True)

        def arg(name):
            if name in args:
                i = args.index(name)
                if i + 1 < len(args):
                    return args[i + 1]
            return None

        bad = []
        for item in (arg("--report") or "").split(";"):
            if not item.strip():
                continue
            p = [x.strip() for x in item.split(",")]
            s, e = (_valid_date(p[1]), _valid_date(p[2])) if len(p) in (3, 4) else (None, None)
            if len(p) in (3, 4) and p[0] in rp.DATE_FIELDS and s and e and s <= e:
                org = p[3] if len(p) == 4 else rp.DEFAULT_ORG
                for a, b in sp.month_pieces(s, min(e, datetime.date.today())):
                    self._add_report_task(org, p[0], a, b)
            else:
                bad.append(item)
        for item in (arg("--transfer") or "").split(";"):
            if not item.strip():
                continue
            p = [x.strip() for x in item.split(",")]
            rng = None
            if len(p) == 3:
                m1, m2 = re.fullmatch(r"(\d{4})-(\d{1,2})", p[1]), re.fullmatch(r"(\d{4})-(\d{1,2})", p[2])
                try:
                    if m1 and m2:
                        rng = month_ranges(int(m1[1]), int(m1[2]), int(m2[1]), int(m2[2]),
                                           bool(self.split_var.get()))
                    elif _valid_date(p[1]) and _valid_date(p[2]):
                        rng = [(_valid_date(p[1]), _valid_date(p[2]))]
                except ValueError:
                    rng = None
            if rng and p[0]:
                for a, b in rng:
                    if a <= b:
                        self._add_transfer_task(p[0], a, b)
            else:
                bad.append(item)
        spec = arg("--sow")
        if spec:
            for item in spec.split(";"):
                if not item.strip():
                    continue
                try:
                    for t in sw.parse_tasks(item):
                        self._add("sow", f"{t['farm']} · {t['stage']}", farm=t["farm"], stage=t["stage"])
                except ValueError:
                    bad.append(item)
        g_codes = []
        if arg("--genetic-codes"):
            try:
                g_codes = gn.codes_from_file(arg("--genetic-codes"))
                self._append_log(f"已从 {arg('--genetic-codes')} 读取国标编码 {len(g_codes)} 个")
            except Exception as e:
                self._append_log(f"读取国标编码文件失败：{e}")
        for item in (arg("--genetic") or "").split(";"):
            if not item.strip():
                continue
            p = [x.strip() for x in re.split(r"[,，]", item)] + ["", ""]
            farm, a, b = p[0], p[1], p[2]
            rng = None
            if not a and not b:
                rng = [("", "")]
            else:
                m1, m2 = re.fullmatch(r"(\d{4})-(\d{1,2})", a), re.fullmatch(r"(\d{4})-(\d{1,2})", b or a)
                try:
                    if m1 and m2:
                        rng = [(x.isoformat(), y.isoformat()) for x, y in
                               month_ranges(int(m1[1]), int(m1[2]), int(m2[1]), int(m2[2]), bool(self.g_split.get()))]
                    elif _valid_date(a) and _valid_date(b or a):
                        rng = [(_valid_date(a).isoformat(), _valid_date(b or a).isoformat())]
                except ValueError:
                    rng = None
            if rng and all(x <= y for x, y in rng):
                farm = "" if farm in ("不限", gn.ALL_FARM_TEXT) else farm
                for x, y in rng:
                    t = gn.make_task(farm, x, y, g_codes)
                    self._add("genetic", gn.task_name(t), trait=t["trait"], farm=farm, start=x, end=y,
                              codes=tuple(g_codes))
            else:
                bad.append(item)
        if arg("--ebv-codes"):
            try:
                e_codes = gn.codes_from_file(arg("--ebv-codes"))
            except Exception as e:
                e_codes = []
                self._append_log(f"读取育种值国标编码文件失败：{e}")
            if e_codes:
                self._append_log(f"已从 {arg('--ebv-codes')} 读取国标编码 {len(e_codes)} 个（育种值查询）")
                self._add("ebv", ev.task_name(ev.make_task(e_codes)), codes=tuple(e_codes))
            else:
                bad.append("--ebv-codes " + arg("--ebv-codes"))
        if arg("--ped-codes"):
            try:
                p_codes = gn.codes_from_file(arg("--ped-codes"))
            except Exception as e:
                p_codes = []
                self._append_log(f"读取系谱查询国标编码文件失败：{e}")
            p_farms = split_names(arg("--ped-farms") or "") or list(self.lists[LIST_PEDIGREE])
            p_status = {"在场": "只查在场", "离场": "只查离场"}.get(arg("--ped-status") or "", pd.STATUS_DEFAULT)
            if p_codes and p_farms:
                t = pd.make_task(p_codes, p_farms, p_status)
                self._add("pedigree", pd.task_name(t), codes=tuple(t["codes"]), farms=tuple(t["farms"]),
                          status=p_status)
            else:
                bad.append("--ped-codes " + arg("--ped-codes") + "（缺少编码或场）")
        for b in bad:
            self._append_log(f"忽略无效的命令行任务：{b}")
        self._refresh_count()
        if "--run-schedule" in args:
            # 由 Windows 计划任务启动：按保存的计划运行，跑完自动退出
            self._append_log(f"由计划任务启动：{sp.describe(self.sched)}")
            self.root.after(800, lambda: self._run_schedule(auto=True) or self._after_auto_run())
            return
        if self.tasks:
            self._append_log(f"已从命令行载入 {len(self.tasks)} 个任务，自动开始 …")
            self.root.after(600, self._start)


def main():
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
