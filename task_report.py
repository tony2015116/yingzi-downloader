# -*- coding: utf-8 -*-
"""
母猪综合查询（统计报表）—— 按日期字段批量导出
================================================================
首页「养殖生产」→「统计报表」→ 报表卡片「母猪综合查询」→ 每组：重置 → 设日期 → 查询 → 导出 → 按条件命名。
查询结果超过平台单次导出上限时自动按日期拆分，各段文件放进以任务命名的文件夹（多级表头，不合并）；
重新下载分段时先清空该文件夹，不会出现 _1 _2 重复文件。文件名带条数，下载目录里有「母猪综合查询下载记录.csv」。
保存位置按日期范围自动分到「2026年 / 06月」子文件夹。

    python task_report.py               # 按 CONFIG["FILTER_GROUPS"] 下载
    python task_report.py --ask         # 交互式输入条件
    python task_report.py --headless    # 后台运行
"""

import os
import re
import sys
import csv
import time
import math
import calendar
import shutil
import datetime
from collections import ChainMap

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    ElementNotInteractableException,
    NoSuchElementException,
    StaleElementReferenceException,
    InvalidSessionIdException,
    WebDriverException,
)

import yz_core as core
from yz_core import (StopRequested, log, wait_until, fmt_secs, sanitize, safe_click, _js, _raise_if_dead,
                     _handles, _current_handle, JS_VIS, JS_CLICK_BY_TEXT, JS_VISIBLE_TEXT, EXPORT_ERR_CSS,
                     net_hook, net_reset, net_state, net_pick, wait_loading_done, click_text, find_menu,
                     switch_to_new_window, confirm_dialog_if_any, page_snapshot, save_debug, open_url,
                     table_state, wait_query_done, query_total, is_limit_msg, watch_dirs, list_files,
                     recent_partials, move_replace, period_folder, xlsx_data_rows, append_csv, rel_to)

TITLE = "母猪综合查询"

# 只写本功能特有的参数；TIMEOUT / POLL / KEEP_LOGIN 等公共参数在 yz_core.CONFIG
CONFIG = ChainMap({
    "REPORT_ENTRIES_URL": "https://bpm.yingzi.com/#/report/entries",   # 菜单点不到时直接打开
    "REPORT_NAME": TITLE,
    # 报表里有两个表格：「母猪生产综合查询（怀孕转出猪只）」和「母猪生产综合查询」，导出的是后者，条数只读它的
    "REPORT_WIDGET": "母猪生产综合查询",
    "DOWNLOAD_DIR": r"D:\test\downloads\母猪综合查询",
    "RESET_BETWEEN_ITERATIONS": True,   # 每组前先点「重置」
    "RANGE_SEPARATOR": " ~ ",
    "DOWNLOAD_TIMEOUT": 600,            # 导出后等待文件的最长时间
    "QUERY_TIMEOUT": 360,               # 点「查询」后最多等多久
    "QUERY_IDLE_GRACE": 12,             # 没有任何请求 / 表格变化时，多久按完成处理
    "EXPORT_RETRY_WAIT": 120,           # 点「导出」后多久没动静就重点一次
    "EXPORT_NO_FILE_WAIT": 25,          # 导出请求结束后多久没文件就判定本次失败
    "EXPORT_ROW_LIMIT": 50000,          # 平台单次导出上限
    "SPLIT_FILL_RATIO": 0.8,            # 拆分时每段目标条数 = 上限 × 比例
    "REFRESH_RECENT_MONTHS": 1,         # 最近几个整月即使已下载过也重新下载（平台上可能还有补录）；0 = 都跳过
    "TRUNCATE_MARGIN": 50,              # 导出文件条数 ≥ 上限 − 此值，视为被平台截断（平台可能把表头也算进 5 万行）
    # 平台限制每次查询只能填一个日期字段。日期："2026-10-01 ~ 2026-10-08" 或 "2026-10-08"
    "FILTER_GROUPS": [
        {"name": "配种日期_2026-10-01~2026-10-08", "org": "广西扬翔农牧有限责任公司",
         "filters": {"配种日期": "2026-10-01 ~ 2026-10-08"}},
    ],
}, core.CONFIG)

DATE_FIELDS = ["首配日期", "配种日期", "配种结果日期", "预产期", "分娩日期", "断奶日期", "离场日期"]
GUI_FIELDS = ["配种日期", "断奶日期", "分娩日期"]

# 报表所在窗口句柄。导出时浏览器可能临时开下载标签页，必须按句柄精确切回报表窗口。
_WIN = {"report": None}
RECORD_NAME = f"{TITLE}下载记录.csv"


class NeedSplit(Exception):
    """查询结果超过单次导出上限，需要拆分日期范围。total 为查询总条数（未知时为 None）。"""

    def __init__(self, total=None, msg=""):
        super().__init__(msg or f"数据超过导出上限（{total or '未知'} 条）")
        self.total = total


DEFAULT_ORG = "广西扬翔农牧有限责任公司"
KEEP_ORG = ("", "当前场", "当前", "当前公司", "默认")


JS_REPORT_TOTAL = JS_VIS + """
const want = arguments[0];
const num = el => { if (!el) return null; const m = (el.textContent || '').replace(/\\s+/g, '').match(/共([\\d,]+)条/);
  return m ? parseInt(m[1].replace(/,/g, ''), 10) : null; };
for (const w of document.querySelectorAll('.react-grid-item')) {
  const h = w.querySelector('h4');
  const title = h ? (h.getAttribute('title') || h.textContent || '').trim() : '';
  if (title !== want) continue;
  const n = num(Array.from(w.querySelectorAll('.davinci-pagination-total-text')).find(vis));
  if (n !== null) return n;
  if (Array.from(w.querySelectorAll('.davinci-table-empty .davinci-empty, .davinci-table-placeholder')).some(vis)) return 0;
  return null;                                   // 找到了表格但还没出条数（加载中）
}
return -1;                                       // 页面上没有这个标题的表格（报表改版）
"""


def query_total(driver):
    """「母猪生产综合查询」表格下方的「共 N 条」。只读这个表格，不会读到上面「怀孕转出猪只」表格的条数。"""
    v = _js(driver, JS_REPORT_TOTAL, CONFIG["REPORT_WIDGET"])
    if v == -1:
        return core.query_total(driver)              # 找不到指定标题时退回通用读法
    try:
        return int(v) if v is not None else None
    except (TypeError, ValueError):
        return None


def make_group(field, start, end, org=""):
    """org：页面右上角要选的公司 / 场；留空 = 不切换，用页面当前选中的。"""
    org = (org or "").strip()
    name = f"{field}_{start}~{end}"
    return {"name": f"{org}_{name}" if org else name, "org": org, "filters": {field: f"{start} ~ {end}"}}


# =====================================================================
# 导航
# =====================================================================
JS_FIND_REPORT_CARD = JS_VIS + """
const name = arguments[0];
for (const lab of document.querySelectorAll("[class*='report-item-label']")) {
  if (norm(lab.textContent) === name && vis(lab)) return lab;
}
return null;
"""
JS_CARD_COUNT = JS_VIS + "return Array.from(document.querySelectorAll(\"[class*='report-item-label']\")).filter(vis).length;"


def _on_report_entries(driver):
    try:
        if "/report/entries" in driver.current_url:
            return True
    except WebDriverException as e:
        _raise_if_dead(e)
    return bool(_js(driver, JS_CARD_COUNT, default=0))


def _open_entries_url(driver):
    open_url(driver, CONFIG["REPORT_ENTRIES_URL"])
    if "user/login" in driver.current_url:
        raise RuntimeError("打开统计报表页面时跳回了登录页（登录已失效）")


def _open_report_entries(driver):
    """进入「统计报表」页面（报表卡片列表），并确认确实进去了。"""
    if _on_report_entries(driver):
        return
    if "bpm.yingzi.com" not in driver.current_url and not find_menu(driver, "统计报表"):
        log("进入「养殖生产」模块 ...")
        before = set(_handles(driver))
        try:
            click_text(driver, "养殖生产", timeout=15)
            switch_to_new_window(driver, before)
            wait_loading_done(driver, timeout=60)
            wait_until(lambda: find_menu(driver, "统计报表"), CONFIG["TIMEOUT"])
        except (StopRequested, InvalidSessionIdException):
            raise
        except Exception as e:
            _raise_if_dead(e)
            log(f"未能通过「养殖生产」进入（{str(e)[:60]}），直接打开统计报表页面")
            _open_entries_url(driver)

    if not _on_report_entries(driver):
        log("进入「统计报表」...")
        menu = wait_until(lambda: find_menu(driver, "统计报表"), 10)
        if menu is not None:
            before = set(_handles(driver))
            safe_click(driver, menu)
            switch_to_new_window(driver, before, wait=1.0)
            if not wait_until(lambda: _on_report_entries(driver), 10):
                log("点击菜单后页面没有跳转，改为直接打开统计报表页面")
                _open_entries_url(driver)
        else:
            log("页面上没有看到「统计报表」菜单，直接打开统计报表页面")
            _open_entries_url(driver)

    if not wait_until(lambda: _on_report_entries(driver), CONFIG["TIMEOUT"]):
        save_debug(driver, "未进入统计报表")
        raise RuntimeError(f"未能进入「统计报表」页面（当前地址：{driver.current_url[:80]}）")
    wait_loading_done(driver, timeout=30)
    log("已进入「统计报表」页面")


def goto_report(driver, max_attempts=3):
    """首页 → 养殖生产 → 统计报表 → 报表卡片「母猪综合查询」（带重试）。"""
    name = CONFIG["REPORT_NAME"]
    for attempt in range(1, max_attempts + 1):
        try:
            try:
                driver.switch_to.default_content()
            except WebDriverException as e:
                _raise_if_dead(e)
            _open_report_entries(driver)

            log(f"点击报表「{name}」...")
            card = wait_until(lambda: _js(driver, JS_FIND_REPORT_CARD, name), CONFIG["TIMEOUT"])
            if card is None:
                save_debug(driver, "未找到报表卡片")
                n = _js(driver, JS_CARD_COUNT, default=0)
                raise RuntimeError(f"「统计报表」页面上未找到报表「{name}」（页面上共 {n} 个报表卡片）")
            before = set(_handles(driver))
            safe_click(driver, card)
            switch_to_new_window(driver, before, wait=1.5)
            wait_loading_done(driver)
            if enter_report_frame(driver):
                _WIN["report"] = _current_handle(driver)
                log(f"已进入「{name}」页面")
                return True
            log("未能进入报表，准备重试导航 ...")
        except (StopRequested, InvalidSessionIdException):
            raise
        except Exception as e:
            _raise_if_dead(e)
            log(f"导航第 {attempt} 次尝试异常: {type(e).__name__}: {str(e)[:100]}")
        time.sleep(2)
    save_debug(driver, "进入报表失败")
    raise RuntimeError(f"多次尝试后仍未能进入「{name}」报表页，请检查网络或页面是否有变化。")


# 页面上可见的 iframe，按面积从大到小；排除客服聊天窗、浏览器插件侧栏等
JS_FRAMES = JS_VIS + """
return Array.from(document.querySelectorAll('iframe'))
  .filter(f => vis(f) && !/chat\\.html|youmind|about:blank/i.test((f.getAttribute('src') || '') + ' ' + (f.id || '')))
  .sort((a, b) => b.offsetWidth * b.offsetHeight - a.offsetWidth * a.offsetHeight);
"""


def _pickers_here(driver):
    return driver.find_elements(By.CSS_SELECTOR, ".davinci-calendar-picker")


def enter_report_frame(driver):
    """切换到承载报表内容的 iframe（DaVinci 报表），并等筛选区渲染出来。"""
    def find_frame():
        for f in driver.find_elements(By.TAG_NAME, "iframe"):
            src = f.get_attribute("src") or ""
            if "davinci" in src or "share.html" in src:
                return f
        return None

    if _pickers_here(driver):          # 报表直接渲染在页面里（没有 iframe）
        net_hook(driver)
        log("已进入报表内容区")
        return True

    target = wait_until(find_frame, 15)
    cands = [target] if target is not None else (_js(driver, JS_FRAMES, default=[]) or [])
    for f in cands:
        try:
            driver.switch_to.frame(f)
        except WebDriverException as e:
            _raise_if_dead(e)
            continue
        net_hook(driver)
        if wait_until(lambda: _pickers_here(driver), CONFIG["TIMEOUT"] if f is target else 10):
            log("已切换到报表内容区")
            return True
        if f is target:
            log("报表筛选区加载超时（继续尝试）")
            return True
        driver.switch_to.default_content()
    log("未找到报表 iframe")
    return False


def ensure_report_frame(driver):
    """确认当前在报表窗口的报表 iframe 内；不在则切回去。报表窗口已不存在时抛错，交给上层重开报表。"""
    rh = _WIN.get("report")
    cur = _current_handle(driver)
    if rh and cur != rh:
        if rh not in _handles(driver):
            raise RuntimeError("报表窗口已被关闭")
        driver.switch_to.window(rh)
    else:
        try:
            if driver.find_elements(By.CSS_SELECTOR, ".davinci-calendar-picker"):
                return
        except WebDriverException as e:
            _raise_if_dead(e)
    try:
        driver.switch_to.default_content()
    except WebDriverException as e:
        _raise_if_dead(e)
    if not enter_report_frame(driver):
        raise RuntimeError("未找到报表 iframe")


def select_report_org(driver, org):
    """报表页右上角的公司 / 场选择框（在报表 iframe 外面）：切到指定公司 / 场，再回到报表 iframe。"""
    org = (org or "").strip()
    if org in KEEP_ORG:
        return
    rh = _WIN.get("report")
    if rh and _current_handle(driver) != rh:
        if rh not in _handles(driver):
            raise RuntimeError("报表窗口已被关闭")
        driver.switch_to.window(rh)
    driver.switch_to.default_content()
    if core.select_org(driver, org):
        # 切换后报表 iframe 会重新加载：等顶层页面请求结束、加载动画消失，再重新进入 iframe
        time.sleep(1.0)
        wait_until(lambda: not (net_state(driver) or {}).get("p") and not core.spinning(driver), 30, poll=0.3)
        driver.switch_to.default_content()
        if not enter_report_frame(driver):
            raise RuntimeError("切换公司 / 场后未能重新进入报表")
        wait_loading_done(driver, 30)
        log(f"已切换到「{org}」")
    else:
        log(f"公司 / 场：当前已是「{org}」")
        ensure_report_frame(driver)


# =====================================================================
# 筛选字段
# =====================================================================
def _label_candidates(driver, label, extra=8):
    # 排除报表表格里的同名表头（第一轮查询后表头会出现「配种日期」等同名列）
    xp = (f"//*[contains(text(),'{label}')][not(ancestor::table)]"
          "[not(ancestor::*[contains(@class,'ant-table') or contains(@class,'el-table')])]")
    return [c for c in driver.find_elements(By.XPATH, xp)
            if c.text.strip() and len(c.text.strip()) <= len(label) + extra]


def _nearest_controls(driver, label, css, timeout=None):
    """根据字段名找到它所在行中包含指定控件的最近容器。返回 (container, controls) 或 (None, None)。"""
    def find():
        cands = _label_candidates(driver, label)
        if not cands:
            return None
        el = min(cands, key=lambda c: len(c.text.strip()))
        for _ in range(8):
            el = el.find_element(By.XPATH, "..")
            controls = el.find_elements(By.CSS_SELECTOR, css)
            if controls:
                return (el, controls)
        return ("none",)          # 字段存在但没有该类控件：立即返回

    r = wait_until(find, timeout or CONFIG["TIMEOUT"])
    if not r or r[0] == "none":
        return None, None
    return r


def _visible_options(driver):
    out = []
    for css in (".ant-select-item-option", ".el-select-dropdown__item", "li[role='option']"):
        out += [o for o in driver.find_elements(By.CSS_SELECTOR, css) if o.is_displayed()]
    return out


def _click_dropdown_option(driver, value):
    opts = _visible_options(driver)
    for o in opts:
        if o.text.strip() == value:
            safe_click(driver, o)
            return True
    for o in opts:
        if value in o.text:
            safe_click(driver, o)
            return True
    return False


def set_text_input(driver, label, value):
    _, controls = _nearest_controls(driver, label, "input, textarea")
    if not controls:
        raise RuntimeError(f"未找到文本输入字段「{label}」")
    inp = controls[0]
    safe_click(driver, inp)
    inp.send_keys(Keys.CONTROL, "a")
    inp.send_keys(value)


def set_select(driver, label, value):
    _, controls = _nearest_controls(driver, label, ".ant-select, .el-select")
    if not controls:
        raise RuntimeError(f"未找到下拉字段「{label}」")
    safe_click(driver, controls[0])
    wait_until(lambda: _visible_options(driver), 3, poll=0.2)
    if _click_dropdown_option(driver, value):
        return
    try:
        search_input = driver.find_element(By.CSS_SELECTOR, ".ant-select-selection-search-input")
        search_input.send_keys(value)
        if wait_until(lambda: _click_dropdown_option(driver, value), 3, poll=0.25):
            return
    except NoSuchElementException:
        pass
    raise RuntimeError(f"下拉字段「{label}」中未找到选项「{value}」")


# ---------------------------------------------------------------------
# DaVinci 日历点选
# ---------------------------------------------------------------------
MONTH_CN = ["一月", "二月", "三月", "四月", "五月", "六月",
            "七月", "八月", "九月", "十月", "十一月", "十二月"]
MONTH_EN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _month_titles(m):
    return [MONTH_CN[m - 1], f"{m}月", MONTH_EN[m - 1], str(m)]


def _parse_month(t):
    t = (t or "").strip()
    d = re.sub(r"\D", "", t)
    if d:
        return int(d)
    if t in MONTH_CN:
        return MONTH_CN.index(t) + 1
    for i, ab in enumerate(MONTH_EN):
        if t.lower().startswith(ab.lower()):
            return i + 1
    return 0


def _raw_click(el):
    """日历弹窗内直接点击（不 scrollIntoView，避免滚动导致弹窗关闭）；点不了就用 JS 点。"""
    try:
        el.click()
    except (ElementClickInterceptedException, StaleElementReferenceException,
            ElementNotInteractableException):
        try:
            el.parent.execute_script("arguments[0].click();", el)
        except Exception:
            pass


def _find_vis(scope, locs, timeout=3.0):
    def f():
        for by, sel in locs:
            for e in scope.find_elements(by, sel):
                if e.is_displayed():
                    return e
        return None
    return wait_until(f, timeout, poll=0.15)


def _panel_ym(part):
    try:
        y = part.find_element(By.CSS_SELECTOR, ".davinci-calendar-year-select").text
        m = part.find_element(By.CSS_SELECTOR, ".davinci-calendar-month-select").text
        return int(re.sub(r"\D", "", y) or 0), _parse_month(m)
    except (NoSuchElementException, StaleElementReferenceException):
        return None


def _navigate_panel_to(driver, part, year, month, max_clicks=40):
    """兜底：用上/下月、上/下年按钮逐级翻到指定年月。"""
    for _ in range(max_clicks):
        ym = _panel_ym(part)
        if ym is None:
            return False
        if ym == (year, month):
            return True
        diff = (year - ym[0]) * 12 + (month - ym[1])
        if diff > 0:
            css = ".davinci-calendar-next-year-btn" if diff >= 12 else ".davinci-calendar-next-month-btn"
        else:
            css = ".davinci-calendar-prev-year-btn" if diff <= -12 else ".davinci-calendar-prev-month-btn"
        btns = part.find_elements(By.CSS_SELECTOR, css)
        if not btns:
            return False
        _raw_click(btns[0])
        wait_until(lambda: _panel_ym(part) != ym, 1.0, poll=0.1)
    return False


def _panel_jump_to(driver, part, year, month):
    """把单个日历面板跳到指定年月：年选择 → 目标年，月选择 → 目标月。"""
    if _panel_ym(part) == (year, month):
        return True
    ysel = _find_vis(part, [(By.CSS_SELECTOR, ".davinci-calendar-year-select")], 2)
    if ysel is None:
        return False
    _raw_click(ysel)

    year_locs = [(By.XPATH, f".//td[@title='{year}']"), (By.XPATH, f".//td[@title='{year}年']"),
                 (By.XPATH, f".//td[normalize-space(.)='{year}']"),
                 (By.XPATH, f".//td[normalize-space(.)='{year}年']")]
    ytd = _find_vis(part, year_locs, 2)
    if ytd is None:
        for _ in range(12):
            heads = part.find_elements(By.CSS_SELECTOR, ".davinci-calendar-year-panel-year-select-content")
            nums = [int(x) for x in re.findall(r"\d{4}", heads[0].text)] if heads else []
            if not nums:
                break
            lo, hi = min(nums), max(nums)
            if lo <= year <= hi:
                break
            css = (".davinci-calendar-year-panel-next-decade-btn, .davinci-calendar-year-panel-next-year-btn"
                   if year > hi else
                   ".davinci-calendar-year-panel-prev-decade-btn, .davinci-calendar-year-panel-prev-year-btn")
            btns = part.find_elements(By.CSS_SELECTOR, css)
            if not btns:
                break
            _raw_click(btns[0])
            time.sleep(0.2)
        ytd = _find_vis(part, year_locs, 1.5)
        if ytd is None:
            return False
    _raw_click(ytd)

    msel = _find_vis(part, [(By.CSS_SELECTOR, ".davinci-calendar-month-select")], 2)
    if msel is None:
        return False
    if _panel_ym(part) == (year, month):
        return True
    _raw_click(msel)
    titles = _month_titles(month)
    mtd = _find_vis(part, [(By.XPATH, f".//td[@title='{t}']") for t in titles]
                    + [(By.XPATH, f".//td[normalize-space(.)='{t}']") for t in titles], 2)
    if mtd is None:
        return False
    _raw_click(mtd)
    return bool(wait_until(lambda: _panel_ym(part) == (year, month), 2, poll=0.15))


def _click_day(part, y, m, d, timeout=3):
    titles = [f"{y}年{m}月{d}日", f"{y}-{m:02d}-{d:02d}", f"{y}/{m}/{d}", f"{y}/{m:02d}/{d:02d}"]
    td = _find_vis(part, [(By.XPATH, f".//td[@title='{t}']") for t in titles], timeout)
    if td is None and _panel_ym(part) == (y, m):
        td = _find_vis(part, [(By.XPATH,
                               ".//td[not(contains(@class,'last-month')) and not(contains(@class,'next-month'))]"
                               f"[normalize-space(.)='{d}']")], 1)
    if td is None:
        return False
    _raw_click(td)
    return True


JS_DATE_PICKERS = """
const labels = arguments[0], out = {};
for (const p of document.querySelectorAll('.davinci-calendar-picker')) {
  if (p.closest('table, .ant-table, .el-table')) continue;
  let el = p;
  for (let i = 0; i < 6 && el.parentElement; i++) {
    el = el.parentElement;
    const c = el.cloneNode(true);
    c.querySelectorAll('.davinci-calendar-picker').forEach(x => x.remove());
    const t = (c.textContent || '').replace(/\\s+/g, '');
    const hits = labels.filter(l => t.includes(l));
    if (hits.length === 1) { if (!(hits[0] in out)) out[hits[0]] = p; break; }
    if (hits.length > 1) break;
  }
}
return out;
"""


def _date_pickers(driver):
    """{字段名: 日历控件}：只包含筛选区的日期框（从控件反查所属字段名，不受表格同名列干扰）。"""
    return _js(driver, JS_DATE_PICKERS, DATE_FIELDS, default={}) or {}


def _field_picker(driver, label):
    p = _date_pickers(driver).get(label)
    if p is not None:
        return p
    try:   # 兜底：从字段名向上找（已排除表格）
        cands = _label_candidates(driver, label)
        if not cands:
            return None
        el = min(cands, key=lambda c: len(c.text.strip()))
        for _ in range(6):
            el = el.find_element(By.XPATH, "..")
            picks = el.find_elements(By.CSS_SELECTOR, ".davinci-calendar-picker")
            if len(picks) == 1:
                return picks[0]
            if len(picks) > 1:
                return None
    except (NoSuchElementException, StaleElementReferenceException):
        pass
    return None


def _picker_values(picker):
    return [i.get_attribute("value") or "" for i in
            picker.find_elements(By.CSS_SELECTOR, ".davinci-calendar-range-picker-input")]


def _clear_picker(driver, picker):
    icons = picker.find_elements(By.CSS_SELECTOR, ".davinci-calendar-picker-clear")
    if not icons:
        return False
    try:
        ActionChains(driver).move_to_element(picker).perform()
    except Exception:
        pass
    _raw_click(icons[0])
    if wait_until(lambda: not any(_picker_values(picker)), 1.5, poll=0.15):
        return True
    # 清除图标只在悬停时显示；后台运行时悬停可能无效，直接用 JS 触发
    driver.execute_script("arguments[0].click();", icons[0])
    return bool(wait_until(lambda: not any(_picker_values(picker)), 1.5, poll=0.15))


def filled_date_fields(driver):
    out = []
    for label, p in _date_pickers(driver).items():
        try:
            if any(_picker_values(p)):
                out.append(label)
        except StaleElementReferenceException:
            pass
    return out


def clear_all_date_fields(driver, keep=None):
    """清空筛选区所有已填的日期字段（keep 除外）。平台限制每次查询只能保留一个日期字段。"""
    cleared, failed = [], []
    for _ in range(2):
        pickers = _date_pickers(driver)
        todo = [(l, p) for l, p in pickers.items() if l != keep]
        pending = False
        for label, picker in todo:
            try:
                if not any(_picker_values(picker)):
                    continue
                pending = True
                if _clear_picker(driver, picker):
                    cleared.append(label)
                    failed = [f for f in failed if f != label]
                elif label not in failed:
                    failed.append(label)
            except InvalidSessionIdException:
                raise
            except StaleElementReferenceException:
                pending = True
            except Exception as e:
                _raise_if_dead(e)
        if not pending:
            break
    if cleared:
        log(f"已清空日期字段: {', '.join(dict.fromkeys(cleared))}")
    if failed:
        log(f"警告：以下日期字段未能清空: {', '.join(failed)}")
    return not failed


def _visible_container(driver, timeout):
    return wait_until(lambda: next(
        (c for c in driver.find_elements(By.CSS_SELECTOR, ".davinci-calendar-picker-container")
         if c.is_displayed()), None), timeout, poll=0.15)


def set_date_range(driver, label, start, end):
    """设置日期区间：打开该字段的日历，分别点选开始/结束日期。"""
    sy, sm, sd = (int(x) for x in start.split("-"))
    ey, em, ed = (int(x) for x in end.split("-"))

    picker = wait_until(lambda: _field_picker(driver, label), 5)
    if picker is None:
        raise RuntimeError(f"未找到日期字段「{label}」")

    container = None
    for _ in range(2):
        safe_click(driver, picker)
        container = _visible_container(driver, 5)
        if container is not None:
            break
    if container is None:
        raise RuntimeError(f"「{label}」日历弹窗未打开")

    try:
        left = container.find_element(By.CSS_SELECTOR, ".davinci-calendar-range-left")
        right = container.find_element(By.CSS_SELECTOR, ".davinci-calendar-range-right")

        if not (_panel_jump_to(driver, left, sy, sm) or _navigate_panel_to(driver, left, sy, sm)):
            raise RuntimeError(f"日历面板未能翻到开始年月 {start[:7]}")
        if not _click_day(left, sy, sm, sd):
            raise RuntimeError(f"未找到开始日期格子 {start}")

        target = next((p for p in (left, right) if _panel_ym(p) == (ey, em)), None)
        if target is None:
            target = right
            if not (_panel_jump_to(driver, right, ey, em) or _navigate_panel_to(driver, right, ey, em)):
                raise RuntimeError(f"日历面板未能翻到结束年月 {end[:7]}")
        if not _click_day(target, ey, em, ed):
            raise RuntimeError(f"未找到结束日期格子 {end}")

        def closed():
            try:
                return not container.is_displayed()
            except StaleElementReferenceException:
                return True
        wait_until(closed, 3, poll=0.15)

        vals = _picker_values(picker)
        digits = [re.sub(r"\D", "", v) for v in vals]
        expect = [start.replace("-", ""), end.replace("-", "")]
        if digits[:2] == expect:
            log(f"「{label}」已设为: {vals}")
        else:
            log(f"警告：「{label}」当前值 {vals}，与期望 {start} ~ {end} 不一致，请核对下载结果")
    except InvalidSessionIdException:
        raise
    except Exception as e:
        _raise_if_dead(e)
        save_debug(driver, f"日期点选失败_{label}")
        raise RuntimeError(f"「{label}」日历点选日期失败: {e}")


def _norm_date(s):
    """2026-1-5 / 2026/01/05 → 2026-01-05（旧版遇到不补零的写法会点错格子）。"""
    nums = [int(x) for x in re.findall(r"\d+", str(s))]
    if len(nums) < 3:
        raise ValueError(f"日期格式不正确：{s}")
    return datetime.date(nums[0], nums[1], nums[2]).isoformat()


def apply_filter(driver, field, value):
    """按字段名设置一个筛选条件。value 为字符串或 (start, end) 元组。"""
    if isinstance(value, (list, tuple)) and len(value) == 2:
        set_date_range(driver, field, _norm_date(value[0]), _norm_date(value[1]))
        return
    if isinstance(value, str) and re.search(r"[~～至]", value):
        parts = re.split(r"[~～至]", value)
        set_date_range(driver, field, _norm_date(parts[0]), _norm_date(parts[1]))
        return
    if isinstance(value, str) and re.fullmatch(r"\d{4}[-/]\d{1,2}[-/]\d{1,2}", value.strip()):
        d = _norm_date(value)
        set_date_range(driver, field, d, d)
        return
    _, controls = _nearest_controls(driver, field, ".ant-select, .el-select")
    if controls:
        set_select(driver, field, value)
        return
    _, controls = _nearest_controls(driver, field, "select")
    if controls:
        Select(controls[0]).select_by_visible_text(value)
        return
    set_text_input(driver, field, value)


# =====================================================================
# 查询 / 导出 / 下载
# =====================================================================
def _dirs():
    return watch_dirs(CONFIG["DOWNLOAD_DIR"])


def list_xlsx():
    return list_files(_dirs(), (".xlsx",))


# 导出时的提示：报表在 iframe 里，提示框可能弹在 iframe 里，也可能弹在外层页面，两处都看。
# 错误 / 警告类之外，再看普通消息框（有的平台用 info 样式提示「超过 5 万条」）；是否算超限仍由 is_limit_msg 判断。
# 不看「确认」类弹窗（确定 / 取消）：正文里「单次不能超过 5 万条，确定导出吗？」只是说明，不是拒绝。
MSG_CSS = EXPORT_ERR_CSS + (", .ant-message-notice, .ant-notification-notice, .ant-modal-confirm-info, "
                            ".el-message, .el-notification")
JS_ALL_FRAMES_TEXT = JS_VIS + """
const css = arguments[0], out = [];
const grab = doc => { try {
  for (const e of doc.querySelectorAll(css)) {
    const v = doc.defaultView;
    const cs = v.getComputedStyle(e);
    if (!(e.offsetWidth || e.offsetHeight) || cs.display === 'none' || cs.visibility === 'hidden') continue;
    const t = (e.innerText || '').trim(); if (t && !out.includes(t)) out.push(t);
  } } catch (err) {} };
grab(document);
let w = window; for (let i = 0; i < 4 && w !== w.parent; i++) { w = w.parent; try { grab(w.document); } catch (err) { break; } }
return out.join(' | ').slice(0, 300);
"""


def export_msg(driver):
    return _js(driver, JS_ALL_FRAMES_TEXT, MSG_CSS, default="") or ""


def click_export(driver, timeout):
    def find():
        for b in driver.find_elements(
                By.XPATH, "//button[contains(translate(normalize-space(.),' ',''),'导出')]"):
            if b.is_displayed() and b.is_enabled() and "disabled" not in (b.get_attribute("class") or ""):
                return b
        return None
    btn = wait_until(find, timeout, poll=0.5)
    if btn is None:
        return False
    safe_click(driver, btn)
    log("已点击「导出」")
    return True


def wait_download(driver, before, timeout, total=None, pre_msg=""):
    """等待新的 xlsx 出现。前 15 秒顺带处理确认弹窗；检测到下载中的临时文件自动延长等待；
    导出请求结束后 EXPORT_NO_FILE_WAIT 秒内既没文件也没新的网络活动，提前判定失败。"""
    t0 = time.time()
    end = t0 + timeout
    no_file_wait = CONFIG["EXPORT_NO_FILE_WAIT"]
    last_log, announced, req_logged = t0, False, False
    net_sig, last_act = None, t0
    seen_msgs = set()
    while time.time() < end:
        core.check_stop()
        now = time.time()
        if now - t0 < 30:
            m = export_msg(driver)
            if m and m != pre_msg and m not in seen_msgs:
                seen_msgs.add(m)
                if is_limit_msg(m):
                    if total is not None and total <= CONFIG["EXPORT_ROW_LIMIT"]:
                        log(f"页面提示「{m}」，但查询结果只有 {total} 条，未超上限，不拆分")
                    else:
                        log(f"导出被拒绝：{m}")
                        _js(driver, JS_CLICK_BY_TEXT, ["确定", "知道了", "关闭", "我知道了"],
                            ".ant-modal-wrap, .el-message-box__wrapper, [role='dialog']", "button")
                        raise NeedSplit(total, m)
                else:
                    log(f"导出时页面提示：{m}")
        if now - t0 < 15:
            confirm_dialog_if_any(driver)
        new = list_xlsx() - before
        if new:
            return max(new, key=os.path.getmtime)
        if recent_partials(_dirs()):
            if not announced:
                log("检测到文件正在下载 ...")
                announced = True
            end = max(end, now + 30)
            last_act = now
        else:
            st = net_state(driver)
            if st:
                sig = (st.get("p"), st.get("s"), st.get("d"))
                if sig != net_sig:
                    net_sig, last_act = sig, now
                if st.get("s", 0) > 0 and st.get("p", 0) == 0:
                    if not req_logged:
                        last = net_pick(driver, "last") or {}
                        code = last.get("s")
                        log(f"导出请求已返回（用时 {fmt_secs(now - t0)}），等待文件下载 ...")
                        if code and code >= 400:
                            log(f"警告：导出接口返回状态 {code}")
                        req_logged = True
                    if now - last_act >= no_file_wait:
                        log(f"导出请求结束 {no_file_wait} 秒仍未收到文件")
                        return None
        if now - last_log >= 20:
            log(f"等待下载文件 ... 已等待 {fmt_secs(now - t0)}")
            last_log = now
        time.sleep(0.4)
    return None


def _close_new_windows(driver, before):
    """关闭导出过程中新开的标签页（不动原有窗口），然后切回报表 iframe。"""
    for h in [h for h in _handles(driver) if h not in before and h != _WIN.get("report")]:
        try:
            driver.switch_to.window(h)
            driver.close()
        except WebDriverException as e:
            _raise_if_dead(e)       # 下载标签页可能已自动关闭，忽略
    ensure_report_frame(driver)


JS_TOTAL_DIAG = JS_VIS + """
const out = [];
for (const e of document.querySelectorAll('*')) {
  if (e.children.length > 3 || !vis(e)) continue;
  const t = (e.textContent || '').replace(/\\s+/g, '');
  if (t.length > 30 || !/共[\\d,]+条/.test(t)) continue;
  if (Array.from(e.children).some(c => /共[\\d,]+条/.test((c.textContent || '').replace(/\\s+/g, '')))) continue;
  let p = e, chain = [];
  for (let i = 0; i < 6 && p; i++, p = p.parentElement) chain.push((p.className && p.className.baseVal === undefined ? p.className : '') .toString().split(' ')[0] || p.tagName.toLowerCase());
  out.push(t + '  ←  ' + chain.join(' < '));
  if (out.length >= 8) break;
}
return out;
"""


def _diagnose_total(driver, name, page_total, rows):
    """页面条数和文件不一致时，把页面上所有「共 N 条」及其所在位置写进日志，并保存现场，方便查清读错了哪个数。"""
    try:
        for line in _js(driver, JS_TOTAL_DIAG, default=[]) or []:
            log(f"  页面上的条数：{line}")
        save_debug(driver, f"条数不一致_{name}")
    except Exception:
        pass


def _settled_total(driver, before_total, before_sig, t0, max_wait=25):
    """查询完成后再确认一次条数：
    连续两次读到同一个数、页面没有加载动画、请求都已返回，才算数；
    如果读到的还是查询前的旧数字（表格也没变），多等一会儿，防止把「重置」后全部数据的条数当成查询结果。"""
    last, same_since = None, None
    end = time.time() + max_wait
    while time.time() < end:
        core.check_stop()
        busy = (net_state(driver) or {}).get("p") or core.spinning(driver)
        n = query_total(driver)
        now = time.time()
        if busy or n != last:
            last, same_since = n, (None if busy else now)
        elif same_since and now - same_since >= 1.0:
            stale = (n is not None and n == before_total and table_state(driver).get("sig") == before_sig)
            if not stale or now - t0 > 20:
                return n
        time.sleep(0.4)
    return last


def download_iteration(driver, group, known_total=None):
    """known_total：已经确定的条数（分段合计没超上限、改为整段重导时传入）——不再因条数或提示拆分。"""
    name = group["name"]
    filters = group.get("filters", {})
    t_start = time.time()
    log(f"---- 开始下载组: {name} ----")

    ensure_report_frame(driver)
    select_report_org(driver, group.get("org"))
    if CONFIG["RESET_BETWEEN_ITERATIONS"]:
        try:
            click_text(driver, "重置", timeout=5)
            wait_loading_done(driver)
        except (StopRequested, InvalidSessionIdException):
            raise
        except Exception as e:
            _raise_if_dead(e)
            log("点击「重置」失败，跳过（继续设置筛选条件）")
    clear_all_date_fields(driver)

    for field, value in filters.items():
        log(f"设置筛选: {field} = {value}")
        apply_filter(driver, field, value)

    # 平台限制只能有一个日期字段：查询前再确认一次，多出来的清掉
    date_keys = [f for f in filters if f in DATE_FIELDS]
    if len(date_keys) == 1:
        extra = [f for f in filled_date_fields(driver) if f != date_keys[0]]
        if extra:
            log(f"发现多余的日期条件 {extra}，清除后再查询")
            if not clear_all_date_fields(driver, keep=date_keys[0]):
                raise RuntimeError(f"无法清除多余的日期条件 {extra}")
        if date_keys[0] not in filled_date_fields(driver):
            raise RuntimeError(f"「{date_keys[0]}」日期未设置成功")

    t_set = time.time()
    net_hook(driver)
    net_reset(driver)
    before_state = table_state(driver)
    before_total = query_total(driver)
    log("点击「查询」...")
    t_click = time.time()
    click_text(driver, "查询")
    wait_query_done(driver, before_state, t_click, CONFIG["QUERY_TIMEOUT"], CONFIG["QUERY_IDLE_GRACE"],
                    err_re=r"日期|只能|不能|失败|错误|异常|请选择|超出|超过")
    page_total = _settled_total(driver, before_total, before_state.get("sig"), t_click, max_wait=8)
    t_query = time.time()
    # 页面上的「共 N 条」只作参考（报表是 DaVinci 组件，之前没识别它的加载动画，读到过查询前的旧数字 7850），
    # 是否拆分、文件名里的条数，都以导出的文件为准。
    log(f"页面显示共 {page_total if page_total is not None else '（未读取到）'} 条"
        + (f"，查询前显示 {before_total} 条" if before_total not in (None, page_total) else ""))
    if known_total is None and page_total is not None and page_total > CONFIG["EXPORT_ROW_LIMIT"]:
        # 超过上限：平台要么拒绝导出、要么只导出前 5 万行（文件看起来正常，数据却不全），所以导出前就拆分。
        # 万一读到的是查询前的旧数字（以前出现过），分段下载完发现合计没超上限，会自动改回整段导出一个文件。
        stale = page_total == before_total and table_state(driver).get("sig") == before_state.get("sig")
        if stale:
            log("提示：页面条数和查询前一样，可能还没刷新；先按超限拆分，分段合计不超上限会自动改回整段导出")
        raise NeedSplit(page_total)

    before = list_xlsx()
    handles_before = set(_handles(driver))
    path = None
    for attempt in range(3):
        net_hook(driver)
        net_reset(driver)
        pre_msg = export_msg(driver)
        if not click_export(driver, CONFIG["QUERY_TIMEOUT"]):
            save_debug(driver, f"导出按钮未就绪_{name}")
            raise RuntimeError(f"「导出」按钮长时间未就绪（查询可能超时），组: {name}")
        path = wait_download(driver, before, CONFIG["EXPORT_RETRY_WAIT"], known_total, pre_msg)
        if path is not None:
            break
        if attempt < 2:
            log(f"未收到下载文件，重试导出（第 {attempt + 2} 次）...")
            _close_new_windows(driver, handles_before)
    if path is None:
        path = wait_download(driver, before, CONFIG["DOWNLOAD_TIMEOUT"], known_total)
    if path is None:
        page_snapshot(driver)
        save_debug(driver, f"下载超时_{name}")
        if known_total is None and page_total is not None and \
                page_total > CONFIG["EXPORT_ROW_LIMIT"] - CONFIG["TRUNCATE_MARGIN"] and split_group(group):
            _close_new_windows(driver, handles_before)
            raise NeedSplit(page_total, f"导出没有生成文件，查询结果 {page_total} 条接近上限，按超限拆分")
        raise RuntimeError(f"等待下载文件超时（组: {name}），请检查导出是否触发。")

    t_dl = time.time()
    rows = xlsx_data_rows(path)
    if rows == 0 and page_total:
        log(f"警告：导出文件读出 0 条，但页面显示 {page_total} 条，文件名改用页面条数，请打开文件核对")
        rows = None
    if rows is not None:
        log(f"导出文件共 {rows} 条数据")
        if page_total is not None and page_total != rows:
            log(f"提示：页面显示 {page_total} 条，导出文件实际 {rows} 条，以文件为准")
            _diagnose_total(driver, name, page_total, rows)
        limit = CONFIG["EXPORT_ROW_LIMIT"]
        if known_total is None and rows >= limit - CONFIG["TRUNCATE_MARGIN"]:
            # 达到或接近上限：多半被平台截断了（平台可能把标题、表头也算在 5 万行里），拆分重下
            try:
                os.remove(path)
            except OSError:
                pass
            _close_new_windows(driver, handles_before)
            raise NeedSplit(page_total if page_total and page_total > rows else None,
                            f"导出文件 {rows} 条，接近单次导出上限 {limit} 条，数据可能被截断")
    else:
        log("提示：未能读取导出文件的行数，文件名中的条数使用页面显示的数字")
    _close_new_windows(driver, handles_before)
    log(f"耗时明细：设置条件 {fmt_secs(t_set - t_start)} · 查询 {fmt_secs(t_query - t_set)} · "
        f"导出下载 {fmt_secs(t_dl - t_query)} · 合计 {fmt_secs(time.time() - t_start)}")
    return path, (rows if rows is not None else page_total), page_total


# =====================================================================
# 超限拆分
# =====================================================================
def _date_filter(group):
    """任务里的日期区间条件：(字段, 开始 date, 结束 date)；没有则 None。"""
    for k, v in group.get("filters", {}).items():
        if k not in DATE_FIELDS:
            continue
        if isinstance(v, (list, tuple)) and len(v) == 2:
            a, b = str(v[0]), str(v[1])
        elif isinstance(v, str) and re.search(r"[~～至]", v):
            a, b = [x.strip() for x in re.split(r"[~～至]", v)[:2]]
        else:
            a = b = str(v).strip()
        return k, datetime.date.fromisoformat(_norm_date(a)), datetime.date.fromisoformat(_norm_date(b))
    return None


def split_group(group, total=None):
    """把任务的日期区间拆成若干段。知道总条数时按比例拆成 N 段，否则对半拆。单日无法再拆返回 []。"""
    df = _date_filter(group)
    if df is None:
        return []
    field, a, b = df
    days = (b - a).days + 1
    if days <= 1:
        return []
    if total:
        per = max(1, int(CONFIG["EXPORT_ROW_LIMIT"] * CONFIG["SPLIT_FILL_RATIO"]))
        n = min(days, max(2, math.ceil(total / per)))
    else:
        n = 2
    parts, start = [], a
    for i in range(n):
        span = days // n + (1 if i < days % n else 0)
        end = start + datetime.timedelta(days=span - 1)
        g = dict(group)
        g["filters"] = dict(group["filters"])
        g["filters"][field] = f"{start.isoformat()} ~ {end.isoformat()}"
        g["name"] = (f"{group['org']}_" if group.get("org") else "") + f"{field}_{start.isoformat()}~{end.isoformat()}"
        parts.append(g)
        start = end + datetime.timedelta(days=1)
    return parts


def _stem(group):
    return f"{TITLE}_{sanitize(group['name'])}"


def _count_text(n):
    return f"{n}条" if n is not None else "条数未知"


def _same_task_files(folder, stem):
    """folder 里属于同一个任务的旧文件：stem.xlsx、stem_123条.xlsx、stem_条数未知.xlsx，以及旧版留下的 _1 _2。"""
    pat = re.compile(re.escape(stem) + r"(_\d+条|_条数未知)?(_\d+)?\.xlsx", re.I)
    try:
        return [os.path.join(folder, f) for f in os.listdir(folder) if pat.fullmatch(f)]
    except OSError:
        return []


def _save(tmp, folder, group, n):
    """删掉同一任务的旧文件 → 按「任务名_N条.xlsx」保存（N = 导出文件的实际数据行数）。返回 (路径, N)。"""
    stem = _stem(group)
    for old in _same_task_files(folder, stem):
        try:
            os.remove(old)
            log(f"已删除同一任务之前下载的文件：{os.path.basename(old)}")
        except OSError:
            pass
    final = move_replace(tmp, folder, f"{stem}_{_count_text(n)}", ".xlsx")
    return final, n


def _download_parts(driver, group, folder, depth, done):
    """分段下载：每段存进 folder；某段仍超限会继续拆。done 收集每段的 {range, total, rows, path}。"""
    try:
        tmp, total, page_total = download_iteration(driver, group)
    except NeedSplit as e:
        parts = split_group(group, e.total)
        if not parts or depth >= 6:
            raise RuntimeError(f"「{group['name']}」{e}，且已无法再按日期拆分（单日数据超过上限）")
        log(f"「{group['name']}」仍超过上限，再拆成 {len(parts)} 段")
        for part in parts:
            _download_parts(driver, part, folder, depth + 1, done)
        return
    path, rows = _save(tmp, folder, group, total)
    fld = _date_filter(group)
    done.append({"range": group["filters"][fld[0]] if fld else group["name"], "total": total,
                 "rows": rows, "page": page_total, "path": path})
    log(f"[分段完成] {group['name']} -> {os.path.basename(path)}")


def _write_record(group, total, rows, parts, path):
    fld = _date_filter(group)
    field, a, b = (fld[0], fld[1].isoformat(), fld[2].isoformat()) if fld else ("", "", "")
    append_csv(os.path.join(CONFIG["DOWNLOAD_DIR"], RECORD_NAME),
               ["下载时间", "公司/场", "日期字段", "开始日期", "结束日期", "页面显示条数（参考）", "文件数据行数", "分段数", "保存位置"],
               [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), group.get("org", ""), field, a, b,
                "" if total is None else total, "" if rows is None else rows, parts,
                rel_to(path, CONFIG["DOWNLOAD_DIR"])])


def field_folder_name(field):
    """日期字段 → 子文件夹名：配种日期 → 配种。"""
    return re.sub(r"日期$", "", field) or field


def report_folder(group):
    """母猪综合查询的保存位置：
       母猪综合查询 › 配种 › 2026年 › 06月 › …  —— 日期范围正好是一个完整的自然月
       母猪综合查询 › 配种 › …                  —— 不足一个月（今天、近 7 天、本月至今……）或跨多个月没拆开
       关闭「按年 / 月自动分文件夹」时，都直接放在 配种 / 分娩 / 断奶 文件夹下。"""
    fld = _date_filter(group)
    if fld is None:
        return CONFIG["DOWNLOAD_DIR"]
    field, a, b = fld
    base = os.path.join(CONFIG["DOWNLOAD_DIR"], field_folder_name(field))
    full_month = (a.day == 1 and a.year == b.year and a.month == b.month
                  and b.day == calendar.monthrange(b.year, b.month)[1])
    if CONFIG.get("DATE_FOLDERS", True) and full_month:
        return os.path.join(base, f"{a.year}年", f"{a.month:02d}月")
    return base


def _is_full_month(a, b):
    return (a.day == 1 and a.year == b.year and a.month == b.month
            and b.day == calendar.monthrange(b.year, b.month)[1])


def existing_download(group):
    """这个任务以前下载好的结果：(路径, 条数) —— 整份文件 任务名_N条.xlsx，或带「分段明细.csv」的完整分段文件夹。
    没有返回 None。只看文件在不在：手动删掉文件，下次就会重新下载。"""
    folder, stem = report_folder(group), _stem(group)
    files = [f for f in _same_task_files(folder, stem) if os.path.getsize(f) > 0]
    if files:
        f = max(files, key=os.path.getmtime)
        m = re.search(r"_(\d+)条\.xlsx$", f)
        return f, (int(m.group(1)) if m else None)
    seg = os.path.join(folder, stem + "_分段")
    detail = os.path.join(seg, "分段明细.csv")
    if os.path.isfile(detail):
        rows = None
        try:
            with open(detail, encoding="utf-8-sig") as fh:
                for r in csv.reader(fh):
                    if r and r[0] == "合计" and len(r) >= 4 and r[3].strip().isdigit():
                        rows = int(r[3])
        except OSError:
            pass
        return seg, rows
    return None


def is_recent_month(a, today=None, n=None):
    """a 所在的月是不是最近 n 个已经过完的月（n = 1 时只有上个月）。"""
    today = today or datetime.date.today()
    n = CONFIG["REFRESH_RECENT_MONTHS"] if n is None else n
    diff = (today.year * 12 + today.month) - (a.year * 12 + a.month)
    return 1 <= diff <= n


def skip_info(group, today=None):
    """整月任务、以前已经下载过、又不是最近要刷新的月份 → 返回 (已有路径, 条数)，可以跳过；否则 None。
    不足一个月的任务（本月至今、近 7 天……）总是重新下载。"""
    fld = _date_filter(group)
    if fld is None or not _is_full_month(fld[1], fld[2]) or is_recent_month(fld[1], today):
        return None
    return existing_download(group)


def cleanup_partials(group):
    """整月下载完成后，删掉配种 / 分娩 / 断奶文件夹下被整月文件完全覆盖的零散文件（本月至今、近 7 天等的旧快照）。
    跨月的零散文件，要它涉及的每个月都已有整月文件才删。返回删除的个数。"""
    fld = _date_filter(group)
    if fld is None:
        return 0
    field = fld[0]
    org = group.get("org", "")
    root = os.path.join(CONFIG["DOWNLOAD_DIR"], field_folder_name(field))
    prefix = f"{TITLE}_" + sanitize(f"{org}_{field}" if org else field) + "_"   # 与 _stem() 的命名一致
    pat = re.compile(re.escape(prefix) + r"(\d{4}-\d{2}-\d{2})~(\d{4}-\d{2}-\d{2})(?:_\d+条|_条数未知)?(?:_\d+)?(\.xlsx|_分段)$")
    try:
        names = os.listdir(root)
    except OSError:
        return 0
    covered = {}
    n = 0
    for name in names:
        m = pat.fullmatch(name)
        if not m:
            continue
        a, b = datetime.date.fromisoformat(m.group(1)), datetime.date.fromisoformat(m.group(2))
        if _is_full_month(a, b) or a > b:
            continue                                  # 整月文件（关闭按年月分文件夹时也在这里）不动
        ok = True
        y, mo = a.year, a.month
        while (y, mo) <= (b.year, b.month):
            key = (y, mo)
            if key not in covered:
                last = calendar.monthrange(y, mo)[1]
                g = make_group(field, f"{y}-{mo:02d}-01", f"{y}-{mo:02d}-{last:02d}", org)
                covered[key] = existing_download(g) is not None
            if not covered[key]:
                ok = False
                break
            y, mo = (y + 1, 1) if mo == 12 else (y, mo + 1)
        if not ok:
            continue
        path = os.path.join(root, name)
        try:
            if os.path.isdir(path):
                shutil.rmtree(path)
            else:
                os.remove(path)
            n += 1
            log(f"已删除被整月文件覆盖的零散文件：{field_folder_name(field)}\\{name}")
        except OSError as e:
            log(f"未能删除零散文件 {name}（{e}），可能正被 Excel 打开，下次再清理")
    return n


def download_group(driver, group):
    """下载一个任务，返回 {"path", "total", "rows", "parts"}。
    未超限：保存为 任务名_N条.xlsx（覆盖同一任务的旧文件）；
    超限：按日期拆分，先清空该任务的「_分段」文件夹，再把各段下载进去，并写「分段明细.csv」。"""
    fld = _date_filter(group)
    base = report_folder(group)
    stem = _stem(group)
    seg_dir = os.path.join(base, stem + "_分段")
    try:
        tmp, total, page_total = download_iteration(driver, group)
    except NeedSplit as e:
        parts = split_group(group, e.total)
        if not parts:
            raise RuntimeError(f"「{group['name']}」{e}，且已无法再按日期拆分（单日数据超过上限）")
        why = f"查询结果 {e.total} 条" if e.total else f"页面提示「{e}」"
        log(f"「{group['name']}」{why}，超过单次导出上限 {CONFIG['EXPORT_ROW_LIMIT']} 条，"
            f"自动拆成 {len(parts)} 段：" + "、".join(p["filters"][fld[0]] for p in parts))
        if os.path.isdir(seg_dir):
            log(f"清空之前下载的分段文件（{len(os.listdir(seg_dir))} 个），重新下载全部分段")
            shutil.rmtree(seg_dir, ignore_errors=True)
            if os.path.isdir(seg_dir):
                raise RuntimeError(f"无法清空 {seg_dir}：里面的文件可能正被 Excel 打开，请关闭后重试")
        os.makedirs(seg_dir, exist_ok=True)
        done = []
        for k, part in enumerate(parts, 1):
            log(f"== 分段 {k}/{len(parts)}：{part['name']} ==")
            core.progress(k - 1, len(parts), "段")
            _download_parts(driver, part, seg_dir, 1, done)
        core.progress(len(parts), len(parts), "段")
        tot = sum(d["total"] for d in done) if all(d["total"] is not None for d in done) else None
        rows = sum(d["rows"] for d in done) if all(d["rows"] is not None for d in done) else None
        real = tot if tot is not None else rows
        if real is not None and real <= CONFIG["EXPORT_ROW_LIMIT"]:
            # 各段加起来其实没超上限（之前读到的条数不准）：整段重新导出成一个文件，删掉分段
            log(f"分段合计只有 {real} 条，没有超过单次导出上限，改为整段重新导出成一个文件 ...")
            tmp, total, page_total = download_iteration(driver, group, known_total=real)
            path, rows1 = _save(tmp, base, group, total)
            shutil.rmtree(seg_dir, ignore_errors=True)
            _write_record(group, page_total, rows1, 1, path)
            log(f"[完成] {group['name']} -> {path}（{total if total is not None else rows1} 条）")
            _after_month(group)
            return {"path": path, "total": total, "rows": rows1, "parts": 1}
        with open(os.path.join(seg_dir, "分段明细.csv"), "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["分段", "日期范围", "页面显示条数（参考）", "文件数据行数", "文件"])
            for k, d in enumerate(done, 1):
                w.writerow([k, d["range"], "" if d.get("page") is None else d["page"],
                            "" if d["rows"] is None else d["rows"], os.path.basename(d["path"])])
            w.writerow(["合计", group["filters"][fld[0]], "", "" if rows is None else rows, ""])
        for old in _same_task_files(base, stem):      # 之前没拆分时下载的整份文件已过时
            try:
                os.remove(old)
            except OSError:
                pass
        pages = [d.get("page") for d in done]
        _write_record(group, sum(pages) if all(p is not None for p in pages) else None, rows, len(done), seg_dir)
        log(f"[完成] {group['name']}：共 {len(done)} 段，{tot if tot is not None else rows} 条 -> {seg_dir}")
        _after_month(group)
        return {"path": seg_dir, "total": tot, "rows": rows, "parts": len(done)}

    path, rows = _save(tmp, base, group, total)
    if os.path.isdir(seg_dir):                       # 之前拆分下载的分段已过时
        shutil.rmtree(seg_dir, ignore_errors=True)
    _write_record(group, page_total, rows, 1, path)
    log(f"[完成] {group['name']} -> {path}（{os.path.getsize(path) / 1024:.0f} KB，"
        f"{total if total is not None else rows} 条）")
    _after_month(group)
    return {"path": path, "total": total, "rows": rows, "parts": 1}


def _after_month(group):
    """整月下载成功后：清理被它覆盖的零散文件。"""
    fld = _date_filter(group)
    if fld and _is_full_month(fld[1], fld[2]):
        try:
            cleanup_partials(group)
        except Exception as e:          # 清理失败不影响下载结果
            log(f"清理零散文件时出错（不影响本次下载）：{e}")


# =====================================================================
# 会话 / 主流程
# =====================================================================
def _goto(driver):
    _WIN["report"] = None
    goto_report(driver)


def open_session(username, password):
    return core.start_session(CONFIG["DOWNLOAD_DIR"], username, password, _goto)


def reload_page(driver):
    core.back_to_home(driver, _WIN.get("report"))
    _goto(driver)


ADAPTER = core.Adapter(goto=_goto, run=download_group, reload=reload_page, name=lambda g: g["name"],
                       download_dir=lambda: CONFIG["DOWNLOAD_DIR"], title=TITLE)


def run_groups(groups, username=None, password=None, status_fn=None):
    return core.run_tasks(groups, ADAPTER, username, password, status_fn, what="下载组")


def ask_groups_interactively():
    print("\n交互模式：平台限制每次查询只能填一个日期字段。")
    fields = {str(i): f for i, f in enumerate(GUI_FIELDS, 1)}
    groups = []
    while True:
        print("\n可选字段: " + "  ".join(f"{k}={v}" for k, v in fields.items()))
        k = input("字段编号（直接回车结束）: ").strip()
        if not k or k not in fields:
            break
        r = input(f"  {fields[k]} 范围（如 2026-10-01 ~ 2026-10-08，或单日 2026-10-08）: ").strip()
        if not r:
            continue
        groups.append({"name": f"{fields[k]}_{sanitize(r)[:24]}", "filters": {fields[k]: r}})
        print(f"已添加: {fields[k]} = {r}")
    if not groups:
        print("未添加任何条件，退出。")
    return groups


def main():
    if "--headless" in sys.argv:
        core.HEADLESS = True
    username, password = core.load_credentials()
    if "--ask" in sys.argv or not CONFIG["FILTER_GROUPS"]:
        groups = ask_groups_interactively()
    else:
        groups = CONFIG["FILTER_GROUPS"]
    if groups:
        run_groups(groups, username, password)


if __name__ == "__main__":
    main()
