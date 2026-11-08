# -*- coding: utf-8 -*-
"""
定时下载 —— 计划的计算部分（不依赖界面，方便测试）
================================================================
计划 = 什么时候跑（每天 / 每周几 / 每月几号 + 时刻）+ 跑哪些任务（任务模板）+ 日期范围规则。
任务模板不带具体日期，每次运行时按规则（昨天、上月……）现算，所以每天 / 每月跑出来的都是新的数据。

运行方式：
  off      不定时
  app      程序窗口开着时到点自动运行
  windows  注册到 Windows「任务计划程序」，到点自动启动本程序运行，跑完自动退出（程序关着也能跑）
"""

import os
import sys
import datetime
import subprocess

TASK_NAME = "影子科技数据下载器_定时下载"
FREQS = {"daily": "每天", "weekly": "每周", "monthly": "每月"}
WEEKDAYS = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
WIN_DAYS = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
MODES = {"off": "不定时", "app": "程序开着时到点自动运行", "windows": "Windows 计划任务（程序关着也能运行）"}
RULES = ["昨天", "近7天（截至昨天）", "本月（截至昨天）", "上周（周一至周日）", "上月", "上季度", "今年（截至昨天）"]
MISS_WINDOW = datetime.timedelta(hours=6)   # 到点时程序在忙 / 刚打开：6 小时内补跑，超过就等下一次

DEFAULT = {"mode": "off", "freq": "daily", "weekday": 0, "monthday": 1, "time": "07:30",
           "rule": "昨天", "templates": [], "last_run": ""}


def normalize(sched):
    out = dict(DEFAULT)
    out.update({k: v for k, v in (sched or {}).items() if k in DEFAULT})
    if out["mode"] not in MODES:
        out["mode"] = "off"
    if out["freq"] not in FREQS:
        out["freq"] = "daily"
    if out["rule"] not in RULES:
        out["rule"] = "昨天"
    try:
        out["weekday"] = min(max(int(out["weekday"]), 0), 6)
        out["monthday"] = min(max(int(out["monthday"]), 1), 28)
    except (TypeError, ValueError):
        out["weekday"], out["monthday"] = 0, 1
    if parse_time(out["time"]) is None:
        out["time"] = "07:30"
    return out


def parse_time(text):
    try:
        h, m = (int(x) for x in str(text).strip().replace("：", ":").split(":"))
        if 0 <= h <= 23 and 0 <= m <= 59:
            return datetime.time(h, m)
    except (TypeError, ValueError):
        pass
    return None


def _month_last(y, m):
    return (datetime.date(y, m, 28) + datetime.timedelta(days=4)).replace(day=1) - datetime.timedelta(days=1)


def date_range(rule, today=None):
    """按规则算出本次要下载的日期范围 (开始, 结束)。"""
    today = today or datetime.date.today()
    y = today - datetime.timedelta(days=1)
    if rule == "昨天":
        return y, y
    if rule == "近7天（截至昨天）":
        return y - datetime.timedelta(days=6), y
    if rule == "本月（截至昨天）":          # 每月 1 号运行时 = 整个上个月
        return y.replace(day=1), y
    if rule == "上周（周一至周日）":
        end = today - datetime.timedelta(days=today.weekday() + 1)
        return end - datetime.timedelta(days=6), end
    if rule == "上月":
        end = today.replace(day=1) - datetime.timedelta(days=1)
        return end.replace(day=1), end
    if rule == "上季度":
        q = (today.month - 1) // 3          # 本季度序号 0..3
        yy, q = (today.year - 1, 3) if q == 0 else (today.year, q - 1)
        return datetime.date(yy, q * 3 + 1, 1), _month_last(yy, q * 3 + 3)
    if rule == "今年（截至昨天）":
        return datetime.date(y.year, 1, 1), y
    raise ValueError(f"未知的日期规则：{rule}")


def split_months(a, b):
    """按自然月切开 [a, b]（首尾不一定是整月）。"""
    out, cur = [], a
    while cur <= b:
        last = min(_month_last(cur.year, cur.month), b)
        out.append((cur, last))
        cur = last + datetime.timedelta(days=1)
    return out


def is_full_month(a, b):
    """[a, b] 是否正好是一个完整的自然月。"""
    return a.day == 1 and a.year == b.year and a.month == b.month and b == _month_last(b.year, b.month)


def month_pieces(a, b):
    """母猪综合查询的日期拆分：范围里包含完整的月，就按自然月拆开（整月一段，月初 / 月末不足一个月的部分各成一段）；
    不包含任何完整月（今天、近 7 天、本月至今、跨月的几天……）则不拆，原样一段。"""
    parts = split_months(a, b)
    if len(parts) > 1 and any(is_full_month(x, y) for x, y in parts):
        return parts
    return [(a, b)]


def _matches(sched, d):
    f = sched["freq"]
    return f == "daily" or (f == "weekly" and d.weekday() == sched["weekday"]) or \
        (f == "monthly" and d.day == sched["monthday"])


def _last_run(sched):
    try:
        return datetime.datetime.fromisoformat(sched.get("last_run") or "")
    except ValueError:
        return None


def is_due(sched, now=None):
    """现在是否该跑：今天是计划日、已过计划时刻（6 小时内）、且这一次还没跑过。"""
    sched = normalize(sched)
    if sched["mode"] != "app" or not sched["templates"]:
        return False
    now = now or datetime.datetime.now()
    if not _matches(sched, now.date()):
        return False
    at = datetime.datetime.combine(now.date(), parse_time(sched["time"]))
    last = _last_run(sched)
    return at <= now < at + MISS_WINDOW and (last is None or last < at)


def next_run(sched, now=None):
    """下一次运行时间（不定时 / 没有任务模板时为 None）。"""
    sched = normalize(sched)
    if sched["mode"] == "off" or not sched["templates"]:
        return None
    now = now or datetime.datetime.now()
    if is_due(sched, now):                  # 已到点、正在等上一批任务做完
        return datetime.datetime.combine(now.date(), parse_time(sched["time"]))
    t = parse_time(sched["time"])
    for k in range(0, 70):
        d = now.date() + datetime.timedelta(days=k)
        at = datetime.datetime.combine(d, t)
        if _matches(sched, d) and at > now:
            return at
    return None


def describe(sched):
    sched = normalize(sched)
    if sched["mode"] == "off":
        return "未设置定时下载"
    when = {"daily": "每天", "weekly": f"每{WEEKDAYS[sched['weekday']]}",
            "monthly": f"每月 {sched['monthday']} 日"}[sched["freq"]]
    return f"{when} {sched['time']} · 下载「{sched['rule']}」· {len(sched['templates'])} 个任务模板"


def template_desc(t):
    if t["kind"] == "report":
        return f"母猪综合查询 · {t.get('org') or '当前公司'} · {t['field']}"
    if t["kind"] == "transfer":
        return f"猪只转移 · {t['farm']}" + ("（按月拆分）" if t.get("split") else "")
    if t["kind"] == "genetic":
        n = len(t.get("codes") or [])
        return (f"遗传性状查询 · {t.get('trait') or '生长性状'} · {t.get('farm') or '全部场'}"
                + (f" · 国标{n}个" if n else "")
                + ("（不限日期）" if not t.get("dated", True) else "（按月拆分）" if t.get("split") else ""))
    if t["kind"] == "ebv":
        return f"育种值查询 · 国标{len(t.get('codes') or [])}个（不按日期）"
    if t["kind"] == "pedigree":
        fs = t.get("farms") or []
        return (f"系谱查询 · 国标{len(t.get('codes') or [])}个 · {len(fs)} 个场 · "
                f"{t.get('status') or '在场优先，再查离场'}（不按日期）")
    return f"母猪档案 · {t['farm']} · {t['stage']}（不按日期）"


def expand(templates, rule, today=None):
    """任务模板 + 日期规则 → 本次要跑的任务：[(kind, params), ...]。母猪档案、育种值查询、系谱查询、不限日期的遗传性状没有日期，原样返回。"""
    a, b = date_range(rule, today)
    out = []
    for t in templates:
        if t["kind"] == "report":
            for x, y in month_pieces(a, b):          # 含完整月的范围（上月以外的季度、本年至今……）按月拆开
                out.append(("report", {"org": t.get("org", ""), "field": t["field"],
                                       "start": x.isoformat(), "end": y.isoformat()}))
        elif t["kind"] == "transfer":
            for x, y in (split_months(a, b) if t.get("split") else [(a, b)]):
                out.append(("transfer", {"farm": t["farm"], "start": x.isoformat(), "end": y.isoformat()}))
        elif t["kind"] == "sow":
            out.append(("sow", {"farm": t["farm"], "stage": t["stage"]}))
        elif t["kind"] == "ebv":
            out.append(("ebv", {"codes": tuple(t.get("codes") or ())}))
        elif t["kind"] == "pedigree":
            out.append(("pedigree", {"codes": tuple(t.get("codes") or ()), "farms": tuple(t.get("farms") or ()),
                                     "status": t.get("status") or "在场优先，再查离场"}))
        elif t["kind"] == "genetic":
            base = {"trait": t.get("trait") or "生长性状", "farm": t.get("farm", ""),
                    "codes": tuple(t.get("codes") or ())}
            if not t.get("dated", True):
                out.append(("genetic", dict(base, start="", end="")))
            else:
                for x, y in (split_months(a, b) if t.get("split") else [(a, b)]):
                    out.append(("genetic", dict(base, start=x.isoformat(), end=y.isoformat())))
    return out


# ---------------------------------------------------------------- Windows 任务计划程序
def launch_command(script=None):
    """计划任务要执行的命令：打包后是 exe 本身，源码运行时是 pythonw + 脚本。"""
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}" --run-schedule'
    py = sys.executable
    w = os.path.join(os.path.dirname(py), "pythonw.exe")
    if os.path.exists(w):
        py = w
    return f'"{py}" "{os.path.abspath(script or sys.argv[0])}" --run-schedule'


def schtasks_create_args(sched, command):
    sched = normalize(sched)
    sc = {"daily": "DAILY", "weekly": "WEEKLY", "monthly": "MONTHLY"}[sched["freq"]]
    args = ["schtasks", "/Create", "/F", "/TN", TASK_NAME, "/TR", command, "/SC", sc,
            "/ST", parse_time(sched["time"]).strftime("%H:%M")]
    if sc == "WEEKLY":
        args += ["/D", WIN_DAYS[sched["weekday"]]]
    elif sc == "MONTHLY":
        args += ["/D", str(sched["monthday"])]
    return args


def _run(args):
    if os.name != "nt":
        return False, "Windows 计划任务只能在 Windows 上设置"
    try:
        r = subprocess.run(args, capture_output=True, creationflags=0x08000000)
        out = (r.stdout or b"") + (r.stderr or b"")
        for enc in ("gbk", "utf-8"):
            try:
                text = out.decode(enc).strip()
                break
            except UnicodeDecodeError:
                continue
        else:
            text = out.decode("utf-8", "replace").strip()
        return r.returncode == 0, text
    except OSError as e:
        return False, str(e)


def register_windows_task(sched, script=None):
    return _run(schtasks_create_args(sched, launch_command(script)))


def delete_windows_task():
    return _run(["schtasks", "/Delete", "/F", "/TN", TASK_NAME])


def windows_task_exists():
    return os.name == "nt" and _run(["schtasks", "/Query", "/TN", TASK_NAME])[0]
