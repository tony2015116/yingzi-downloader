# -*- coding: utf-8 -*-
"""
遗传性状查询（基因组选种平台）—— 按「测定猪场 × 结测日期 × 国标编码」批量导出
================================================================
首页「基因组选种」→ 左侧「性状信息」→「遗传性状查询」→ 标签「生长性状」→ 每个任务：
重置 → 展开全部条件 →（国标编码）点「批量」粘贴 →（结测日期）选起止日期 →（测定猪场）下拉选中
→ 查询 → 右上角「导出」→ 按条件命名、放进年 / 月文件夹。

  · 三个条件都可以留空，留空 = 不限；
  · 国标编码超过 400 个（平台「批量」一次最多 400 个）：每 400 个一批分别导出，最后合并成一个文件；
  · 查询结果超过 5 万条（平台单次导出上限）：按结测日期自动拆分，分段文件放进「…_分段」文件夹；
  · 文件名：遗传性状_生长性状_{测定猪场|全部场}_{结测开始~结束|不限日期}[_国标N个]_{N}条.xlsx
  · 每个任务追加一行到「遗传性状查询下载记录.csv」。

    python task_genetic.py --tasks "南沙母猪场16线,2026-08-01,2026-08-31"   # 场可留空：",2026-08-01,2026-08-31"
    python task_genetic.py --tasks "南沙母猪场16线,," --codes 国标编码.txt
    python task_genetic.py --headless
"""

import os
import re
import sys
import csv
import math
import time
import shutil
import datetime
from collections import ChainMap

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import (
    InvalidSessionIdException,
    WebDriverException,
    NoSuchElementException,
    StaleElementReferenceException,
)

import yz_core as core
from yz_core import (StopRequested, log, wait_until, fmt_secs, sanitize, safe_click, _js, _raise_if_dead,
                     _handles, _current_handle, _n, _digits, JS_VIS, JS_SET_VALUE, JS_VISIBLE_TEXT,
                     EXPORT_ERR_CSS, net_hook, net_reset, net_state, net_pick, net_dump, wait_loading_done,
                     click_text, find_menu, switch_to_new_window, confirm_dialog_if_any, page_snapshot,
                     save_debug, open_url, table_state, wait_query_done, query_total, is_limit_msg,
                     watch_dirs, list_files, recent_partials, move_replace, period_folder, rel_to,
                     xlsx_data_rows, append_csv, check_stop)

TITLE = "遗传性状查询"

CONFIG = ChainMap({
    "GSM_URL": "https://gsm.yingzi.com/#/measure-record/geneticCharacter",   # 菜单点不到时直接打开
    "GSM_HOST": "gsm.yingzi.com",
    "PORTAL_TILE": "基因组选种",           # 登录后首页上要点的入口（不是「养殖生产」）
    "DOWNLOAD_DIR": r"D:\test\downloads\遗传性状查询",
    "QUERY_TIMEOUT": 360,                 # 点「查询」后最多等多久
    "QUERY_IDLE_GRACE": 12,
    "EXPORT_ENABLE_WAIT": 15,             # 查询完成后「导出」按钮最多等多久变为可用
    "EXPORT_RETRY_WAIT": 120,             # 点「导出」后多久没动静就重点一次
    "EXPORT_NO_FILE_WAIT": 25,
    "DOWNLOAD_TIMEOUT": 600,
    "EXPORT_ROW_LIMIT": 50000,            # 平台单次导出上限
    "SPLIT_FILL_RATIO": 0.8,              # 拆分时每段目标条数 = 上限 × 比例
    "CODE_BATCH": 400,                    # 「批量」输入国标编码一次最多几个
    "SPLIT_START": "2015-01-01",          # 没填结测日期又超过上限时，从这一天到今天按日期拆分
    "TASKS": [
        {"trait": "生长性状", "farm": "南沙母猪场16线", "start": "2026-08-01", "end": "2026-08-31", "codes": []},
    ],
}, core.CONFIG)

# 每种性状在页面上的标签、日期框、猪场下拉（目前界面只开放「生长性状」，其他标签以后加这里即可）
TRAITS = {
    "生长性状": {"key": "born", "date_id": "finalBusinessDate", "date_label": "结测",
                 "farm_id": "qp-finalOrgId-eq", "farm_label": "测定猪场"},
}
DEFAULT_TRAIT = "生长性状"
FILE_PREFIX = "遗传性状"
RECORD_NAME = f"{TITLE}下载记录.csv"
DOWNLOAD_EXTS = (".xlsx", ".xls", ".csv", ".zip")
ALL_FARM_TEXT = "全部场"
NO_DATE_TEXT = "不限日期"

_MAIN = {"handle": None}


class NeedSplit(Exception):
    """查询结果超过单次导出上限，需要按日期拆分。total 为页面条数（未知时为 None）。"""

    def __init__(self, total=None, msg=""):
        super().__init__(msg or f"数据超过导出上限（{total if total is not None else '未知'} 条）")
        self.total = total


# =====================================================================
# 国标编码
# =====================================================================
CODE_RE = re.compile(r"[A-Za-z0-9][A-Za-z0-9_\-]{5,39}")


def parse_codes(text):
    """粘贴的文字 → (去重后的编码列表, 不像编码的内容)。换行、空格、逗号、分号、顿号都当分隔符。"""
    codes, bad, seen = [], [], set()
    for tok in re.split(r"[\s,，;；、|]+", str(text or "")):
        tok = tok.strip().strip("'\"")
        if not tok:
            continue
        if not CODE_RE.fullmatch(tok):
            bad.append(tok)
            continue
        key = tok.upper()
        if key not in seen:
            seen.add(key)
            codes.append(tok)
    return codes, bad


def codes_from_file(path):
    """从 txt / csv / xlsx 读取国标编码。Excel：优先读表头含「国标」的那一列，没有就读所有像编码的单元格。"""
    low = path.lower()
    cells = []
    if low.endswith((".xlsx", ".xlsm")):
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb.worksheets[0]
            ws.reset_dimensions()
            rows = list(ws.iter_rows(values_only=True))
        finally:
            wb.close()
        col = hdr = None
        for i, r in enumerate(rows[:5]):
            for j, v in enumerate(r):
                if v is not None and "国标" in str(v):
                    col, hdr = j, i
                    break
            if col is not None:
                break
        if col is not None:
            cells = [r[col] for r in rows[hdr + 1:] if col < len(r)]
        else:
            cells = [v for r in rows for v in r]
        text = "\n".join(str(v) for v in cells if v not in (None, ""))
        codes, _bad = parse_codes(text)
        if col is None:
            codes = [c for c in codes if re.search(r"[A-Za-z]", c) and re.search(r"\d", c)]
        return codes
    if low.endswith(".xls"):
        import xlrd
        book = xlrd.open_workbook(path)
        sh = book.sheet_by_index(0)
        text = "\n".join(str(sh.cell_value(r, c)) for r in range(sh.nrows) for c in range(sh.ncols))
        codes, _bad = parse_codes(text)
        return [c for c in codes if re.search(r"[A-Za-z]", c) and re.search(r"\d", c)]
    for enc in ("utf-8-sig", "gbk", "utf-16"):
        try:
            with open(path, encoding=enc) as f:
                text = f.read()
            break
        except (UnicodeDecodeError, UnicodeError):
            continue
    else:
        raise ValueError("无法识别文件编码，请另存为 UTF-8 的 txt 或 xlsx")
    codes, _bad = parse_codes(text)
    return codes


def code_batches(codes, size=None):
    size = size or CONFIG["CODE_BATCH"]
    codes = list(codes or [])
    return [codes[i:i + size] for i in range(0, len(codes), size)] or [None]


# =====================================================================
# 任务 / 命名
# =====================================================================
def _date(v):
    v = str(v or "").strip()
    if not v:
        return None
    return datetime.date.fromisoformat(v)


def norm_task(task):
    t = dict(task)
    t["trait"] = t.get("trait") or DEFAULT_TRAIT
    if t["trait"] not in TRAITS:
        raise ValueError(f"不支持的性状：{t['trait']}（可选：{' / '.join(TRAITS)}）")
    t["farm"] = (t.get("farm") or "").strip()
    if t["farm"] in (ALL_FARM_TEXT, "不限", "全部"):
        t["farm"] = ""
    a, b = _date(t.get("start")), _date(t.get("end"))
    if (a is None) != (b is None):
        a = b = a or b
    if a and b and a > b:
        a, b = b, a
    t["start"], t["end"] = (a.isoformat() if a else ""), (b.isoformat() if b else "")
    t["codes"] = list(t.get("codes") or [])
    return t


def date_text(task):
    lab = TRAITS[task.get("trait") or DEFAULT_TRAIT]["date_label"]
    return f"{lab}{task['start']}~{task['end']}" if task.get("start") else NO_DATE_TEXT


def make_stem(task):
    task = norm_task(task)
    parts = [FILE_PREFIX, task["trait"], task["farm"] or ALL_FARM_TEXT, date_text(task)]
    if task["codes"]:
        parts.append(f"国标{len(task['codes'])}个")
    return sanitize("_".join(parts))


def task_name(task):
    task = norm_task(task)
    s = f"{task['trait']} · {task['farm'] or ALL_FARM_TEXT} · {date_text(task)}"
    return s + (f" · 国标 {len(task['codes'])} 个" if task["codes"] else "")


def describe_conditions(task):
    task = norm_task(task)
    return (f"测定猪场={task['farm'] or '不限'}，结测日期={task['start'] + '~' + task['end'] if task['start'] else '不限'}，"
            f"国标编码={len(task['codes']) if task['codes'] else '不限'}")


# =====================================================================
# 页面内 JS
# =====================================================================
JS_TABS_PRESENT = JS_VIS + """
const ts = Array.from(document.querySelectorAll('.ant-tabs-tab')).filter(vis).map(t => norm(t.textContent));
return ts.includes('基本性状') && ts.some(t => /性状$/.test(t) && t !== '基本性状');
"""

JS_TRAIT_TAB = JS_VIS + """
const text = arguments[0], key = arguments[1];
for (const t of document.querySelectorAll('.ant-tabs-tab')) {
  if (!vis(t)) continue;
  const btn = t.querySelector('.ant-tabs-tab-btn') || t;
  if (norm(btn.textContent) === text || (btn.id || '').endsWith('-tab-' + key) ||
      (t.getAttribute('data-node-key') || '') === key)
    return {el: btn, active: t.classList.contains('ant-tabs-tab-active')};
}
return null;
"""

# 查询区域：遗传性状查询每个性状标签有自己的一套查询条件（id 会重复，必须限定在当前标签页内找）；
# 育种值查询的标签页是空的，查询表单在标签下方单独的卡片里 —— 这时取页面上那个可见的查询表单所在的卡片。
JS_PANE = JS_VIS + """
const panes = Array.from(document.querySelectorAll('.ant-tabs-tabpane-active')).filter(vis);
const withForm = panes.find(p => Array.from(p.querySelectorAll('form')).some(vis));
if (withForm) return withForm;
const forms = Array.from(document.querySelectorAll('form')).filter(f => vis(f)
  && !f.closest('.ant-modal-wrap, .ant-drawer, .ant-tabs-tabpane:not(.ant-tabs-tabpane-active)'));
const f = forms.find(x => /查\\s*询/.test(x.textContent || '')) || forms[0];
if (f) return f.closest('.ant-card') || f.closest('.yz-search-section') || f;
// 表单还没渲染出来（页面加载中）：不返回空的标签页，让调用方继续等，否则会在空白区域里找不到按钮 / 输入框
return panes.find(p => Array.from(p.querySelectorAll('input, button')).some(vis)) || null;
"""

JS_BATCH_MODAL = JS_VIS + """
const wraps = Array.from(document.querySelectorAll('.ant-modal-wrap')).filter(w => vis(w) && w.querySelector('textarea'));
wraps.sort((a, b) => (/国标|批量/.test((b.querySelector('.ant-modal-title') || {}).textContent || '') ? 1 : 0)
                   - (/国标|批量/.test((a.querySelector('.ant-modal-title') || {}).textContent || '') ? 1 : 0));
for (const w of wraps) {
  const ta = w.querySelector('textarea');
  const btns = Array.from(w.querySelectorAll('.ant-modal-footer button')).filter(vis);
  const ok = btns.find(b => b.classList.contains('ant-btn-primary')) || btns.find(b => /确定|确认/.test(norm(b.textContent)));
  const m = norm(w.querySelector('.ant-modal-body') ? w.querySelector('.ant-modal-body').innerText : '').match(/(\\d+)\\/(\\d+)/);
  return {wrap: w, ta: ta, ok: ok || null, n: m ? parseInt(m[1], 10) : null, max: m ? parseInt(m[2], 10) : null};
}
return null;
"""

JS_PICKER = JS_VIS + """
const want = arguments[0];
const dd = Array.from(document.querySelectorAll('.ant-picker-dropdown'))
  .filter(d => vis(d) && !d.classList.contains('ant-picker-dropdown-hidden'));
if (!dd.length) return null;
const d = dd[dd.length - 1];
const cell = Array.from(d.querySelectorAll('td.ant-picker-cell-in-view')).find(c => c.getAttribute('title') === want && vis(c));
if (cell) return {cell: cell.querySelector('.ant-picker-cell-inner') || cell,
                  disabled: cell.classList.contains('ant-picker-cell-disabled')};
const h = Array.from(d.querySelectorAll('.ant-picker-header')).find(vis);
if (!h) return {err: 'nohead'};
const yb = h.querySelector('.ant-picker-year-btn'), mb = h.querySelector('.ant-picker-month-btn');
const btn = c => Array.from(d.querySelectorAll(c)).find(vis) || null;
return {y: yb ? yb.textContent : '', m: mb ? mb.textContent : '', view: (h.textContent || '').trim().slice(0, 40),
        sp: btn('.ant-picker-header-super-prev-btn'), p: btn('.ant-picker-header-prev-btn'),
        n: btn('.ant-picker-header-next-btn'), sn: btn('.ant-picker-header-super-next-btn')};
"""

JS_PICKER_OK = JS_VIS + """
const b = Array.from(document.querySelectorAll('.ant-picker-dropdown:not(.ant-picker-dropdown-hidden) .ant-picker-ok button')).find(vis);
if (b && !b.disabled) { b.click(); return true; } return false;
"""

JS_PICKER_OPEN = JS_VIS + """
return Array.from(document.querySelectorAll('.ant-picker-dropdown'))
  .some(d => vis(d) && !d.classList.contains('ant-picker-dropdown-hidden'));
"""

JS_SELECT_OPTIONS = JS_VIS + """
const id = arguments[0], name = norm(arguments[1]), exact = arguments[2];
let dds = Array.from(document.querySelectorAll('.ant-select-dropdown'))
  .filter(d => vis(d) && !d.classList.contains('ant-select-dropdown-hidden'));
const own = dds.filter(d => d.querySelector('[id="' + id + '_list"]'));
if (own.length) dds = own;
const out = [];
for (const d of dds) for (const o of d.querySelectorAll('.ant-select-item-option')) {
  if (!vis(o)) continue;
  const t = norm(o.getAttribute('title') || o.textContent);
  if (exact ? t === name : t.includes(name)) out.push(o);
}
return out;
"""

JS_SELECT_TEXTS = JS_VIS + """
const out = [];
for (const d of document.querySelectorAll('.ant-select-dropdown:not(.ant-select-dropdown-hidden)'))
  for (const o of d.querySelectorAll('.ant-select-item-option, .ant-select-item-empty'))
    if (vis(o)) { const t = norm(o.getAttribute('title') || o.textContent); if (t && !out.includes(t)) out.push(t); }
return out.slice(0, 30);
"""

JS_EXPORT_BTN = JS_VIS + """
for (const b of document.querySelectorAll('button')) {
  if (norm(b.textContent) !== '导出' || !vis(b) || b.closest('.ant-modal-wrap, .ant-drawer')) continue;
  const wrap = b.parentElement && b.parentElement.classList.contains('ant-tooltip-disabled-compatible-wrapper') ? b.parentElement : null;
  return {el: b, wrap: wrap, ok: !(b.disabled || b.classList.contains('ant-btn-disabled'))};
}
return null;
"""

JS_TOOLTIP = JS_VIS + """
return Array.from(document.querySelectorAll('.ant-tooltip:not(.ant-tooltip-hidden) .ant-tooltip-inner, .ant-popover:not(.ant-popover-hidden) .ant-popover-inner-content'))
  .filter(vis).map(e => (e.innerText || '').trim()).filter(Boolean).join(' | ').slice(0, 200);
"""

INFO_MSG_CSS = ".ant-message-notice, .ant-notification-notice"
MONTH_EN = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]


def _escape(driver):
    try:
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    except Exception:
        pass


def _click(driver, el):
    try:
        ActionChains(driver).move_to_element(el).pause(0.05).click(el).perform()
    except WebDriverException as e:
        _raise_if_dead(e)
        safe_click(driver, el)


# =====================================================================
# 导航：首页「基因组选种」→ 性状信息 → 遗传性状查询 → 生长性状
# =====================================================================
def _url(driver):
    try:
        return driver.current_url or ""
    except WebDriverException as e:
        _raise_if_dead(e)
        return ""


def on_page(driver):
    return CONFIG["GSM_HOST"] in _url(driver) and bool(_js(driver, JS_TABS_PRESENT, default=False))


def _login_lost(driver):
    u = _url(driver)
    return "user/login" in u or ("account.yingzi.com" in u and CONFIG["GSM_HOST"] not in u)


def goto_page(driver, max_attempts=3):
    for attempt in range(1, max_attempts + 1):
        try:
            try:
                driver.switch_to.default_content()
            except WebDriverException as e:
                _raise_if_dead(e)
            if not on_page(driver):
                if CONFIG["GSM_HOST"] not in _url(driver):
                    log(f"进入「{CONFIG['PORTAL_TILE']}」...")
                    before = set(_handles(driver))
                    try:
                        click_text(driver, CONFIG["PORTAL_TILE"], timeout=15)
                        switch_to_new_window(driver, before, wait=4)
                        wait_loading_done(driver, timeout=60)
                        wait_until(lambda: CONFIG["GSM_HOST"] in _url(driver), 20)
                    except (StopRequested, InvalidSessionIdException):
                        raise
                    except Exception as e:
                        _raise_if_dead(e)
                        log(f"未能通过首页「{CONFIG['PORTAL_TILE']}」进入（{str(e)[:60]}），直接打开遗传性状查询页面")
                        open_url(driver, CONFIG["GSM_URL"])
                if CONFIG["GSM_HOST"] in _url(driver) and not on_page(driver):
                    item = wait_until(lambda: find_menu(driver, "遗传性状查询"), 3)
                    if item is None:
                        grp = wait_until(lambda: find_menu(driver, "性状信息"), 8)
                        if grp is not None:
                            log("展开左侧「性状信息」...")
                            safe_click(driver, grp)
                            item = wait_until(lambda: find_menu(driver, "遗传性状查询"), 5)
                    if item is not None:
                        log("打开「遗传性状查询」...")
                        safe_click(driver, item)
                    if not wait_until(lambda: on_page(driver), 15):
                        log("菜单点击后未看到遗传性状查询页面，直接打开页面地址")
                        open_url(driver, CONFIG["GSM_URL"])
            if _login_lost(driver):
                raise RuntimeError("打开基因组选种平台时跳回了登录页（登录已失效或没有该平台权限）")
            if wait_until(lambda: on_page(driver), CONFIG["TIMEOUT"]):
                wait_loading_done(driver, timeout=30)
                net_hook(driver)
                _MAIN["handle"] = _current_handle(driver)
                log("已进入「遗传性状查询」页面")
                return True
            log("未能进入遗传性状查询页面，准备重试 ...")
        except (StopRequested, InvalidSessionIdException):
            raise
        except Exception as e:
            _raise_if_dead(e)
            log(f"导航第 {attempt} 次尝试异常：{type(e).__name__}: {str(e)[:100]}")
        time.sleep(2)
    save_debug(driver, "进入遗传性状查询失败")
    raise RuntimeError("多次尝试后仍未能进入「遗传性状查询」页面，请确认账号有「基因组选种」平台的权限。")


def ensure_page(driver):
    mh = _MAIN.get("handle")
    if mh and _current_handle(driver) != mh and mh in _handles(driver):
        driver.switch_to.window(mh)
    if not on_page(driver):
        goto_page(driver)


def _qstate(driver):
    """查询完成判断用的页面状态：去掉「导出按钮是否可用」（超过 5 万条时导出按钮一直不可用，不能当作还在加载）。"""
    s = dict(table_state(driver) or {})
    s.pop("exp", None)
    return s


def select_trait(driver, trait):
    tr = TRAITS[trait]
    info = wait_until(lambda: _js(driver, JS_TRAIT_TAB, trait, tr["key"]), 15)
    if not info:
        save_debug(driver, "未找到性状标签")
        raise RuntimeError(f"页面上未找到「{trait}」标签")
    if info.get("active"):
        return False
    log(f"切换到「{trait}」")
    net_hook(driver)
    net_reset(driver)
    before = _qstate(driver)
    safe_click(driver, info["el"])
    if not wait_until(lambda: (_js(driver, JS_TRAIT_TAB, trait, tr["key"]) or {}).get("active"), 6, poll=0.2):
        raise RuntimeError(f"点击「{trait}」后标签未切换")
    wait_query_done(driver, before, time.time(), CONFIG["QUERY_TIMEOUT"], 6, state_fn=_qstate, err_re=None,
                    what="打开标签")
    return True


def pane(driver):
    p = wait_until(lambda: _js(driver, JS_PANE), 10)
    if p is None:
        raise RuntimeError("未找到当前标签页的查询区域")
    return p


def _find_in(driver, el, css):
    try:
        return [e for e in el.find_elements(By.CSS_SELECTOR, css) if e.is_displayed()]
    except StaleElementReferenceException:
        return []
    except WebDriverException as e:
        _raise_if_dead(e)
        return []


def _pane_button(driver, text):
    def find():
        for b in _find_in(driver, pane(driver), "button"):
            if _n(b.text) == text and b.is_enabled():
                return b
        return None
    return wait_until(find, 10)


# =====================================================================
# 查询条件
# =====================================================================
def reset_form(driver):
    b = _pane_button(driver, "重置")
    if b is None:
        log("未找到「重置」按钮，跳过")
        return
    net_hook(driver)
    net_reset(driver)
    before = _qstate(driver)
    safe_click(driver, b)
    wait_query_done(driver, before, time.time(), CONFIG["QUERY_TIMEOUT"], 4, state_fn=_qstate, err_re=None,
                    what="重置")


def expand_form(driver, trait):
    """结测日期、测定猪场在折叠区域里：没显示就点右侧 ︽ 展开。"""
    tr = TRAITS[trait]
    shown = lambda: bool(_find_in(driver, pane(driver), f"[id='{tr['date_id']}']"))
    if shown():
        return
    btns = _find_in(driver, pane(driver), ".more-btn button")
    if not btns:
        raise RuntimeError("未找到展开更多条件的按钮")
    log("展开全部查询条件")
    safe_click(driver, btns[0])
    if not wait_until(shown, 5, poll=0.2):
        raise RuntimeError(f"展开后仍看不到「{tr['date_label']}日期」输入框")


def set_codes(driver, codes, input_id="qp-gbCode-in"):
    """点「批量」→ 粘贴国标编码（每行一个）→ 核对弹窗计数 → 确定。codes 为空时只确认输入框是空的。
    input_id：编码输入框的 id（遗传性状查询 qp-gbCode-in，育种值查询 qp-personId-in）。"""
    if codes and len(codes) > CONFIG["CODE_BATCH"]:
        raise ValueError(f"一批国标编码不能超过 {CONFIG['CODE_BATCH']} 个")
    # 页面可能还在加载：等到能找到「批量」为止。优先用 id 找输入框；id 对不上时，取带「批量」的那个输入框。
    found = wait_until(lambda: _code_input(driver, input_id), 15 if codes else 3)
    inp, trig = found if found else (None, None)
    if not codes:
        if inp is not None and (inp.get_attribute("value") or "").strip():
            driver.execute_script(JS_SET_VALUE, inp, "")
        return
    if trig is None:
        save_debug(driver, "未找到国标编码批量")
        raise RuntimeError("页面上没有找到国标编码输入框旁边的「批量」，页面可能还没加载完或页面结构有变化"
                           "（debug 文件夹里有当时的截图）")
    if inp is not None and (inp.get_attribute("value") or "").strip():
        _clear_input(driver, inp)        # 先清掉上一批的编码，避免弹窗里新旧编码叠加
    safe_click(driver, trig)
    md = wait_until(lambda: _js(driver, JS_BATCH_MODAL), 8)
    if not md or md.get("ta") is None:
        save_debug(driver, "批量国标编码弹窗未出现")
        raise RuntimeError("点「批量」后没有出现「批量输入国标编码」弹窗")
    time.sleep(0.4)             # 等弹窗打开动画结束
    md = _js(driver, JS_BATCH_MODAL) or md
    ta = md["ta"]
    text = "\n".join(codes)
    driver.execute_script(JS_SET_VALUE, ta, "")
    driver.execute_script(JS_SET_VALUE, ta, text)
    if (ta.get_attribute("value") or "").strip() != text:
        ta.click()
        ta.send_keys(Keys.CONTROL, "a")
        ta.send_keys(Keys.BACKSPACE)
        driver.execute_script(JS_SET_VALUE, ta, text)
    md = wait_until(lambda: (lambda m: m if m and m.get("n") == len(codes) else None)(_js(driver, JS_BATCH_MODAL)),
                    3, poll=0.2) or _js(driver, JS_BATCH_MODAL) or {}
    n = md.get("n")
    if n is not None and n != len(codes):
        log(f"警告：批量弹窗显示 {n} 个，本批实际 {len(codes)} 个（平台可能自动去重），继续")
    else:
        log(f"已粘贴国标编码 {len(codes)} 个" + (f"（弹窗显示 {n}/{md.get('max')}）" if n is not None else ""))
    if md.get("ok") is None:
        raise RuntimeError("批量输入弹窗里没有找到「确定」按钮")
    closed = False
    for k in range(3):          # 弹窗刚打开还在动画中时，点击可能没生效：没关就再点一次
        safe_click(driver, md["ok"])
        if wait_until(lambda: not _js(driver, JS_BATCH_MODAL), 4, poll=0.2):
            closed = True
            break
        msg = _js(driver, JS_VISIBLE_TEXT, INFO_MSG_CSS, default="") or ""
        if re.search(r"最多|超过|不能|错误|失败", msg):
            raise RuntimeError(f"批量输入国标编码被拒绝：{msg}")
        md = _js(driver, JS_BATCH_MODAL) or md
        if md.get("ok") is None:
            break
        log("批量输入弹窗未关闭，再点一次「确定」")
    if not closed:
        save_debug(driver, "批量输入弹窗未关闭")
        raise RuntimeError("点「确定」后批量输入弹窗没有关闭")
    if inp is not None:
        v = inp.get_attribute("value") or ""
        if not v.strip():
            log("提示：国标编码输入框里看不到编码（平台可能只在内部记录），以查询结果为准")
        else:
            got = len([x for x in re.split(r"[\s,，;；]+", v) if x])
            if got > len(codes):
                save_debug(driver, "国标编码叠加")
                raise RuntimeError(f"国标编码输入框里有 {got} 个编码，本批只有 {len(codes)} 个（上一批的编码没有清掉）")


def _batch_in(driver, scope):
    for e in _find_in(driver, scope, ".ant-input-group-addon span, .ant-input-group-addon"):
        if _n(e.text) == "批量":
            return e
    return None


def _code_input(driver, input_id):
    """返回 (编码输入框, 「批量」按钮)；找不到「批量」返回 None。"""
    try:
        p = pane(driver)
    except RuntimeError:
        return None
    inp = next(iter(_find_in(driver, p, f"input[id='{input_id}']")), None)
    if inp is not None:
        try:
            grp = inp.find_element(By.XPATH, "./ancestor::span[contains(@class,'ant-input-group-wrapper')][1]")
            trig = _batch_in(driver, grp)
        except WebDriverException as e:
            _raise_if_dead(e)
            trig = None
        if trig is not None:
            return inp, trig
    trig = _batch_in(driver, p)
    if trig is None:
        return None
    try:
        grp = trig.find_element(By.XPATH, "./ancestor::span[contains(@class,'ant-input-group-wrapper')][1]")
        inp = next(iter(_find_in(driver, grp, "input")), inp)
    except WebDriverException as e:
        _raise_if_dead(e)
    return inp, trig


def _clear_input(driver, inp):
    """清空编码输入框：优先点输入框右侧的 ×（和手动操作一样会通知页面），不行再用脚本清空。"""
    try:
        ActionChains(driver).move_to_element(inp).pause(0.2).perform()
        icons = inp.find_elements(By.XPATH, "./ancestor::span[contains(@class,'ant-input-affix-wrapper')][1]"
                                            "//span[contains(@class,'ant-input-clear-icon')]")
        if icons:
            driver.execute_script("arguments[0].click();", icons[0])
    except WebDriverException as e:
        _raise_if_dead(e)
    if wait_until(lambda: not (inp.get_attribute("value") or "").strip(), 2, poll=0.2):
        return
    driver.execute_script(JS_SET_VALUE, inp, "")


def _ym(info):
    y = int(_digits(info.get("y")) or 0)
    m_txt = (info.get("m") or "").strip().lower()
    m = int(_digits(m_txt) or 0)
    if not m:
        m = next((i + 1 for i, k in enumerate(MONTH_EN) if m_txt.startswith(k)), 0)
    return y, m


def _pick_day(driver, day):
    """日期弹出层已打开时，翻到目标月份并点击那一天。"""
    want = day.isoformat()
    for _ in range(120):
        check_stop()
        info = _js(driver, JS_PICKER, want)
        if info is None:
            return False                      # 弹出层没打开
        if info.get("cell") is not None:
            if info.get("disabled"):
                raise RuntimeError(f"日期 {want} 在页面上不可选（可能超出平台允许的范围）")
            _click(driver, info["cell"])
            return True
        if info.get("err"):
            return False
        y, m = _ym(info)
        if not y or not m:
            return False
        diff = (day.year * 12 + day.month) - (y * 12 + m)
        if diff < 0:
            btn = info.get("sp") if diff <= -12 and info.get("sp") else info.get("p")
        else:
            btn = info.get("sn") if diff >= 13 and info.get("sn") else info.get("n")
        if btn is None:
            return False
        driver.execute_script("arguments[0].click();", btn)
        time.sleep(0.08)
    return False


def _picker_inputs(driver, trait):
    p = pane(driver)
    start = next(iter(_find_in(driver, p, f"input[id='{TRAITS[trait]['date_id']}']")), None)
    if start is None:
        raise RuntimeError(f"未找到「{TRAITS[trait]['date_label']}开始日期」输入框")
    box = start.find_element(By.XPATH, "./ancestor::div[contains(concat(' ', normalize-space(@class), ' '), ' ant-picker-range ')][1]")
    ins = box.find_elements(By.CSS_SELECTOR, "input")
    return box, ins[0], ins[-1]


def _date_values(driver, trait):
    _box, a, b = _picker_inputs(driver, trait)
    return _digits(a.get_attribute("value")), _digits(b.get_attribute("value"))


def set_date_range(driver, trait, start, end):
    lab = TRAITS[trait]["date_label"]
    want = (_digits(start.isoformat()), _digits(end.isoformat()))
    if _date_values(driver, trait) == want:
        return
    box, a_in, b_in = _picker_inputs(driver, trait)
    ok = False
    for attempt in range(2):
        _escape(driver)
        safe_click(driver, a_in)
        if not wait_until(lambda: _js(driver, JS_PICKER_OPEN, default=False), 4, poll=0.2):
            continue
        if _pick_day(driver, start):
            time.sleep(0.2)
            if not _js(driver, JS_PICKER_OPEN, default=False):     # 选完开始日期弹层就关了：再点结束框
                safe_click(driver, b_in)
                wait_until(lambda: _js(driver, JS_PICKER_OPEN, default=False), 3, poll=0.2)
            if _pick_day(driver, end):
                _js(driver, JS_PICKER_OK)
                ok = wait_until(lambda: _date_values(driver, trait) == want, 3, poll=0.2)
        if ok:
            break
        log(f"{lab}日期第 {attempt + 1} 次点选未成功，重试 ...")
    if not ok:
        # 兜底：去掉只读属性直接输入日期并回车
        log(f"改用直接输入的方式填写{lab}日期")
        for el, d in ((a_in, start), (b_in, end)):
            driver.execute_script("arguments[0].removeAttribute('readonly');", el)
            safe_click(driver, el)
            el.send_keys(Keys.CONTROL, "a")
            el.send_keys(d.isoformat())
            el.send_keys(Keys.ENTER)
            time.sleep(0.3)
        ok = wait_until(lambda: _date_values(driver, trait) == want, 3, poll=0.2)
    _escape(driver)
    if not ok:
        save_debug(driver, f"{lab}日期设置失败")
        raise RuntimeError(f"{lab}日期未能设置为 {start} ~ {end}（当前显示 {_date_values(driver, trait)}）")
    log(f"已设置{lab}日期：{start} ~ {end}")


def _farm_select(driver, trait):
    fid = TRAITS[trait]["farm_id"]
    inp = next(iter(_find_in(driver, pane(driver), f"input[id='{fid}']")), None)
    if inp is None:
        # 只读的 antd 下拉框输入框本身可能不可见：按 id 找，再取外层
        els = pane(driver).find_elements(By.CSS_SELECTOR, f"input[id='{fid}']")
        inp = els[0] if els else None
    if inp is None:
        raise RuntimeError(f"未找到「{TRAITS[trait]['farm_label']}」下拉框")
    sel = inp.find_element(By.XPATH, "./ancestor::div[contains(concat(' ', normalize-space(@class), ' '), ' ant-select ')][1]")
    return sel, inp


def current_farm(driver, trait):
    sel, _inp = _farm_select(driver, trait)
    items = sel.find_elements(By.CSS_SELECTOR, ".ant-select-selection-item")
    if not items:
        return ""
    return (items[0].get_attribute("title") or items[0].text or "").strip()


# 下拉打开后，在下拉里逐屏滚动查找选项（antd 的虚拟列表只渲染看得见的十来个选项）
JS_DD_SCROLL = JS_VIS + """
const id = arguments[0], step = arguments[1];
let dds = Array.from(document.querySelectorAll('.ant-select-dropdown'))
  .filter(d => vis(d) && !d.classList.contains('ant-select-dropdown-hidden'));
const own = dds.filter(d => d.querySelector('[id="' + id + '_list"]'));
if (own.length) dds = own;
const d = dds[0]; if (!d) return {open: false};
const h = d.querySelector('.rc-virtual-list-holder') || d.querySelector('.ant-select-dropdown .rc-virtual-list') || d;
if (step === 'top') { h.scrollTop = 0; return {open: true, end: false}; }
if (step) h.scrollTop = h.scrollTop + Math.max(40, h.clientHeight * 0.8);
return {open: true, end: h.scrollTop + h.clientHeight >= h.scrollHeight - 2};
"""


def _find_in_dropdown(driver, list_id, name):
    """在已打开的下拉里找文字完全一致的选项：先看当前屏，再从头逐屏往下滚。返回选项元素或 None。"""
    st = _js(driver, JS_DD_SCROLL, list_id, None) or {}
    if not st.get("open"):
        return None
    _js(driver, JS_DD_SCROLL, list_id, "top")
    for _ in range(200):
        time.sleep(0.15)
        opts = _js(driver, JS_SELECT_OPTIONS, list_id, name, True, default=[]) or []
        if opts:
            return opts[0]
        st = _js(driver, JS_DD_SCROLL, list_id, None) or {}
        if st.get("end") or not st.get("open"):
            return None
        _js(driver, JS_DD_SCROLL, list_id, True)
    return None


def _all_dropdown_texts(driver, list_id):
    """把下拉里的全部选项文字读出来（逐屏滚动），报错时提示用户真实的名称。"""
    seen = []
    _js(driver, JS_DD_SCROLL, list_id, "top")
    for _ in range(200):
        time.sleep(0.15)
        for t in _js(driver, JS_SELECT_TEXTS, default=[]) or []:
            if t not in seen and t != "暂无数据":
                seen.append(t)
        st = _js(driver, JS_DD_SCROLL, list_id, None) or {}
        if st.get("end") or not st.get("open"):
            break
        _js(driver, JS_DD_SCROLL, list_id, True)
    return seen


def set_farm(driver, trait, farm):
    lab, fid = TRAITS[trait]["farm_label"], TRAITS[trait]["farm_id"]
    cur = current_farm(driver, trait)
    if not farm:
        if cur:
            sel, _inp = _farm_select(driver, trait)
            ActionChains(driver).move_to_element(sel).perform()
            clr = sel.find_elements(By.CSS_SELECTOR, ".ant-select-clear")
            if clr:
                driver.execute_script("arguments[0].dispatchEvent(new MouseEvent('mousedown', {bubbles: true}));"
                                      "arguments[0].click();", clr[0])
            if current_farm(driver, trait):
                raise RuntimeError(f"「{lab}」未能清空（当前为 {current_farm(driver, trait)}）")
        return
    if _n(cur) == _n(farm):
        log(f"「{lab}」已是 {farm}")
        return
    sel, inp = _farm_select(driver, trait)
    safe_click(driver, sel.find_element(By.CSS_SELECTOR, ".ant-select-selector"))
    try:
        inp.send_keys(Keys.CONTROL, "a")
        inp.send_keys(farm)
    except WebDriverException as e:
        _raise_if_dead(e)
    opts = wait_until(lambda: _js(driver, JS_SELECT_OPTIONS, fid, farm, True, default=[]), 6, poll=0.3)
    if not opts:
        part = _js(driver, JS_SELECT_OPTIONS, fid, farm, False, default=[]) or []
        if len(part) == 1:
            opts = part
    seen = None
    if not opts:
        try:                                     # 清掉搜索文字（有的下拉按内部编号搜索，输入中文名搜不到）
            inp.send_keys(Keys.CONTROL, "a")
            inp.send_keys(Keys.BACKSPACE)
            time.sleep(0.6)
        except WebDriverException as e:
            _raise_if_dead(e)
        o = _find_in_dropdown(driver, fid, farm)           # 不输入，逐屏滚动查找
        if o is not None:
            opts = [o]
        else:
            seen = _all_dropdown_texts(driver, fid)
    if not opts:
        seen = seen or [t for t in (_js(driver, JS_SELECT_TEXTS, default=[]) or []) if t != "暂无数据"]
        save_debug(driver, f"未找到{lab}_{farm}")
        _escape(driver)
        raise RuntimeError(f"「{lab}」下拉列表里没有找到「{farm}」。列表中可见的名称："
                           f"{'、'.join(seen[:20]) or '（无）'}（请在「编辑场列表」里改成页面上的名称）")
    try:
        opts[0].click()
    except WebDriverException as e:
        _raise_if_dead(e)
        driver.execute_script("arguments[0].click();", opts[0])
    if not wait_until(lambda: _n(current_farm(driver, trait)) == _n(farm) or
                      (len(opts) == 1 and current_farm(driver, trait)), 5, poll=0.2):
        save_debug(driver, f"{lab}选择失败_{farm}")
        raise RuntimeError(f"点选「{farm}」后「{lab}」显示为「{current_farm(driver, trait)}」，选择未生效")
    _escape(driver)
    log(f"已选择{lab}：{current_farm(driver, trait)}")


def _settled_total(driver, before_total, before_sig, t0, max_wait=25):
    """查询完成后再确认一次「共 N 条」：连续读到同一个数、没有加载动画才算数；还是查询前的旧数字就多等一会。"""
    last, same_since = object(), None
    end = time.time() + max_wait
    while time.time() < end:
        check_stop()
        busy = (net_state(driver) or {}).get("p") or core.spinning(driver)
        n = query_total(driver)
        now = time.time()
        if busy or n != last:
            last, same_since = n, (None if busy else now)
        elif same_since and now - same_since >= 1.0:
            stale = n is not None and n == before_total and _qstate(driver).get("sig") == before_sig
            if not stale or now - t0 > 15:
                return n
        time.sleep(0.4)
    return last if isinstance(last, int) else None


def run_query(driver):
    b = _pane_button(driver, "查询")
    if b is None:
        raise RuntimeError("未找到「查询」按钮")
    net_hook(driver)
    net_reset(driver)
    before = _qstate(driver)
    before_total = query_total(driver)
    t0 = time.time()
    log("点击「查询」...")
    safe_click(driver, b)
    wait_query_done(driver, before, t0, CONFIG["QUERY_TIMEOUT"], CONFIG["QUERY_IDLE_GRACE"], state_fn=_qstate,
                    err_re=r"失败|错误|异常|请选择|请输入|不能|只能|超出|超过")
    total = _settled_total(driver, before_total, before.get("sig"), t0)
    if total is None and _qstate(driver).get("empty"):
        total = 0                       # 没有数据时平台不显示分页，也就没有「共 N 条」
    return total


# =====================================================================
# 导出 / 下载
# =====================================================================
def _dirs():
    return watch_dirs(CONFIG["DOWNLOAD_DIR"])


def list_downloads():
    return list_files(_dirs(), DOWNLOAD_EXTS)


def export_ready(driver, total=None):
    """等「导出」按钮可用。返回 (按钮, 不可用原因)。"""
    info = wait_until(lambda: (lambda b: b if b and b.get("ok") else None)(_js(driver, JS_EXPORT_BTN)),
                      CONFIG["EXPORT_ENABLE_WAIT"], poll=0.4)
    if info:
        return info["el"], ""
    info = _js(driver, JS_EXPORT_BTN)
    if not info:
        return None, "页面右上角没有找到「导出」按钮"
    tip = ""
    target = info.get("wrap") or info.get("el")
    try:
        ActionChains(driver).move_to_element(target).perform()
        tip = wait_until(lambda: _js(driver, JS_TOOLTIP, default=""), 3, poll=0.3) or ""
        ActionChains(driver).move_by_offset(0, 120).perform()
    except WebDriverException as e:
        _raise_if_dead(e)
    return None, tip or "「导出」按钮一直是灰色的"


def wait_download(driver, before, timeout, total=None, dirs=None):
    """等待新文件出现；处理确认弹窗、限额提示、接口直接返回文件地址等情况。返回文件路径或 None。
    dirs：要监控的下载目录（默认 = 本功能的下载目录）。"""
    dirs = dirs or _dirs()
    t0 = time.time()
    end = t0 + timeout
    last_log, announced, req_logged = t0, False, False
    net_sig, last_act = None, t0
    seen_msgs, opened = set(), set()
    while time.time() < end:
        check_stop()
        now = time.time()
        if now - t0 < 60:
            m = _js(driver, JS_VISIBLE_TEXT, EXPORT_ERR_CSS + ", " + INFO_MSG_CSS, default="") or ""
            if m and m not in seen_msgs:
                seen_msgs.add(m)
                if is_limit_msg(m) and (total is None or total > CONFIG["EXPORT_ROW_LIMIT"]):
                    raise NeedSplit(total, f"导出被拒绝：{m}")
                if re.search(r"失败|错误|异常|无权限|没有权限", m):
                    raise RuntimeError(f"导出失败，页面提示：{m}")
                log(f"导出时页面提示：{m}")
                if re.search(r"下载中心|任务列表|后台|稍后|生成中|排队", m):
                    log("提示：平台在后台生成文件，继续等待；若一直没有文件，请把这条提示发给维护人员")
                    end = max(end, now + CONFIG["DOWNLOAD_TIMEOUT"])
        if now - t0 < 15:
            confirm_dialog_if_any(driver)
        new = list_files(dirs, DOWNLOAD_EXTS) - before
        if new:
            return max(new, key=os.path.getmtime)
        if recent_partials(dirs):
            if not announced:
                log("检测到文件正在下载 ...")
                announced = True
            end = max(end, now + 30)
            last_act = now
        else:
            dump = net_dump(driver)
            for u in dump.get("urls") or []:
                if u not in opened and now - t0 > 4:
                    opened.add(u)
                    log(f"接口返回了文件地址，直接打开下载：{u[:120]}")
                    try:
                        driver.execute_script("window.open(arguments[0], '_blank');", u)
                    except WebDriverException as e:
                        _raise_if_dead(e)
                    end = max(end, now + 120)
                    last_act = now
            st = net_state(driver)
            if st:
                sig = (st.get("p"), st.get("s"), st.get("d"))
                if sig != net_sig:
                    net_sig, last_act = sig, now
                if st.get("s", 0) > 0 and st.get("p", 0) == 0:
                    if not req_logged:
                        last = net_pick(driver, "last") or {}
                        log(f"导出请求已返回（用时 {fmt_secs(now - t0)}），等待文件下载 ...")
                        if last.get("s") and last["s"] >= 400:
                            log(f"警告：导出接口返回状态 {last['s']}")
                        req_logged = True
                    if now - last_act >= CONFIG["EXPORT_NO_FILE_WAIT"]:
                        log(f"导出请求结束 {CONFIG['EXPORT_NO_FILE_WAIT']} 秒仍未收到文件")
                        return None
        if now - last_log >= 20:
            log(f"等待下载文件 ... 已等待 {fmt_secs(now - t0)}")
            last_log = now
        time.sleep(0.4)
    return None


def _close_new_windows(driver, before):
    main = _MAIN.get("handle")
    for h in _handles(driver):
        if h not in before and h != main:
            try:
                driver.switch_to.window(h)
                driver.close()
            except WebDriverException as e:
                _raise_if_dead(e)
    hs = _handles(driver)
    if main in hs:
        driver.switch_to.window(main)
    elif hs:
        driver.switch_to.window(hs[0])
        _MAIN["handle"] = hs[0]


def export_once(driver, cond):
    """按一组条件（场 / 日期 / 一批编码）查询并导出一次。
    返回 (临时文件, 条数, 页面条数)；0 条返回 None；超过上限抛 NeedSplit。"""
    trait = cond["trait"]
    ensure_page(driver)
    _escape(driver)
    select_trait(driver, trait)
    reset_form(driver)
    expand_form(driver, trait)
    set_codes(driver, cond.get("codes") or [])
    if cond.get("start"):
        set_date_range(driver, trait, _date(cond["start"]), _date(cond["end"]))
    elif any(_date_values(driver, trait)):
        raise RuntimeError(f"「重置」后{TRAITS[trait]['date_label']}日期仍有值，无法按「不限日期」查询")
    set_farm(driver, trait, cond.get("farm") or "")
    t_set = time.time()

    total = run_query(driver)
    log(f"查询结果：共 {total if total is not None else '（未读取到）'} 条")
    if total == 0:
        return None
    if total is not None and total > CONFIG["EXPORT_ROW_LIMIT"]:
        raise NeedSplit(total, f"查询结果 {total} 条，超过单次导出上限 {CONFIG['EXPORT_ROW_LIMIT']} 条")

    before = list_downloads()
    handles_before = set(_handles(driver))
    path = None
    for attempt in range(3):
        btn, why = export_ready(driver, total)
        if btn is None:
            if is_limit_msg(why) or re.search(r"5\s*万|50000|超过|上限", why):
                raise NeedSplit(total, f"导出按钮不可用：{why}")
            save_debug(driver, "导出按钮不可用")
            raise RuntimeError(f"「导出」不可用：{why}")
        net_hook(driver)
        net_reset(driver)
        safe_click(driver, btn)
        log("已点击「导出」" + (f"（第 {attempt + 1} 次）" if attempt else ""))
        path = wait_download(driver, before, CONFIG["EXPORT_RETRY_WAIT"], total)
        if path:
            break
        _close_new_windows(driver, handles_before)
        _escape(driver)
        log("未收到下载文件，重试导出 ...")
    if path is None:
        path = wait_download(driver, before, CONFIG["DOWNLOAD_TIMEOUT"], total)
    if path is None:
        for r in net_dump(driver).get("last") or []:
            log(f"  最近请求：{r.get('u')} → {r.get('s')} {r.get('ct') or ''} {r.get('cd') or ''}")
        page_snapshot(driver)
        save_debug(driver, "下载超时_遗传性状")
        raise RuntimeError("等待下载文件超时，请检查导出是否触发（debug 文件夹里有当时的截图）")
    _close_new_windows(driver, handles_before)
    t_dl = time.time()

    rows = xlsx_data_rows(path)
    if rows == 0 and total:
        log(f"警告：导出文件读出 0 条，但页面显示 {total} 条，条数改用页面数字，请打开文件核对")
        rows = None
    if rows is not None:
        if total is not None and rows != total:
            log(f"提示：页面显示 {total} 条，导出文件实际 {rows} 条，以文件为准")
        if rows >= CONFIG["EXPORT_ROW_LIMIT"]:
            try:
                os.remove(path)
            except OSError:
                pass
            raise NeedSplit(None, f"导出文件正好 {rows} 条（达到单次导出上限，数据可能被截断）")
    log(f"导出完成：{rows if rows is not None else total} 条（导出用时 {fmt_secs(t_dl - t_set)}）")
    return path, (rows if rows is not None else total), total


def split_range(a, b, total=None):
    """按日期把 [a, b] 拆成几段：知道总条数时按比例拆，否则对半拆。单日无法再拆返回 []。"""
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
        parts.append((start, end))
        start = end + datetime.timedelta(days=1)
    return parts


def _collect(driver, cond, work, pieces, ctx, depth=0):
    """导出一组条件；超过上限时按结测日期拆分（没填日期就从 SPLIT_START 到今天拆），结果收集到 pieces。"""
    try:
        res = export_once(driver, cond)
    except NeedSplit as e:
        lab = TRAITS[cond["trait"]]["date_label"]
        if cond.get("start"):
            a, b = _date(cond["start"]), _date(cond["end"])
        else:
            a, b = _date(CONFIG["SPLIT_START"]), datetime.date.today()
            ctx.setdefault("undated_total", e.total)
            log(f"没有填{lab}日期但数据超过上限：改为按{lab}日期 {a} ~ {b} 分段导出"
                f"（{lab}日期为空或早于 {a} 的记录导不出来，结束时会核对条数）")
        parts = split_range(a, b, e.total)
        if not parts or depth >= 8:
            raise RuntimeError(f"{e}，且已无法再按{lab}日期拆分（单日数据超过上限）")
        log(f"{e}：按{lab}日期拆成 {len(parts)} 段：" + "、".join(f"{x}~{y}" for x, y in parts))
        for x, y in parts:
            log(f"== 分段 {x} ~ {y} ==")
            _collect(driver, dict(cond, start=x.isoformat(), end=y.isoformat()), work, pieces, ctx, depth + 1)
        return
    if res is None:
        log("本段没有数据")
        return
    tmp, n, page = res
    dst = move_replace(tmp, work, f"{len(pieces) + 1:03d}", os.path.splitext(tmp)[1].lower() or ".xlsx")
    pieces.append({"path": dst, "rows": n, "page": page, "start": cond.get("start", ""),
                   "end": cond.get("end", ""), "batch": cond.get("batch")})


# =====================================================================
# 合并 / 保存
# =====================================================================
def _header_end(path):
    """导出表格的表头占几行（标题行 + 多级表头，规则同 yz_core.xlsx_data_rows）。"""
    import openpyxl
    merges = core._xlsx_merges(path)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        ws.reset_dimensions()
        head = []
        for r in ws.iter_rows(values_only=True):
            head.append(r)
            if len(head) >= 2:
                break
    finally:
        wb.close()
    filled = lambda r: sum(1 for v in r if v not in (None, ""))
    start = 2 if len(head) >= 2 and filled(head[0]) == 1 and filled(head[1]) >= 3 else 1
    return max([b for a, b in merges if a == start] + [start])


def merge_xlsx(paths, out):
    """把几个结构相同的导出文件合并成一个：表头取第一个文件，后面的只追加数据行。返回合并后的数据行数。"""
    import openpyxl
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    wb_out = openpyxl.Workbook(write_only=True)
    ws_out = None
    rows, header, renum = 0, None, False
    bold, fill = Font(bold=True), PatternFill("solid", fgColor="F2F4F7")
    for i, p in enumerate(paths):
        he = _header_end(p)
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        try:
            ws = wb.worksheets[0]
            ws.reset_dimensions()
            if ws_out is None:
                ws_out = wb_out.create_sheet(ws.title or "Sheet1")
            for r_i, r in enumerate(ws.iter_rows(values_only=True), 1):
                if r_i <= he:
                    if i == 0:
                        if r_i == he:
                            header = [v for v in r]
                            renum = bool(header) and _n(header[0]) in ("序号", "#", "No.", "NO")
                            for c, v in enumerate(header, 1):
                                ws_out.column_dimensions[get_column_letter(c)].width = \
                                    max(10, min(40, len(str(v or "")) * 2 + 2))
                            try:
                                ws_out.freeze_panes = f"A{he + 1}"
                            except Exception:
                                pass
                        cells = []
                        for v in r:
                            c = WriteOnlyCell(ws_out, value=v)
                            c.font, c.fill, c.alignment = bold, fill, Alignment(horizontal="center", vertical="center")
                            cells.append(c)
                        ws_out.append(cells)
                    elif r_i == he and header is not None:
                        trim = lambda x: [str(v).strip() if v is not None else "" for v in x]
                        if trim(r)[:len(header)] != trim(header)[:len(r)]:
                            log(f"警告：第 {i + 1} 个文件的表头与第 1 个不一致，仍按列顺序合并，请核对")
                    continue
                if not any(v not in (None, "") for v in r):
                    continue
                r = list(r)
                rows += 1
                if renum and r:
                    r[0] = rows                 # 「序号」各批都从 1 开始，合并后重新连续编号
                ws_out.append(r)
        finally:
            wb.close()
    if ws_out is None:
        ws_out = wb_out.create_sheet("Sheet1")
    wb_out.save(out)
    return rows


def _same_task_files(folder, stem):
    pat = re.compile(re.escape(stem) + r"(_\d+条|_条数未知)?(_\d+)?\.(xlsx|xls|csv|zip)", re.I)
    try:
        return [os.path.join(folder, f) for f in os.listdir(folder) if pat.fullmatch(f)]
    except OSError:
        return []


def _count_text(n):
    return f"{n}条" if n is not None else "条数未知"


def _rmtree(path):
    if os.path.isdir(path):
        shutil.rmtree(path, ignore_errors=True)
        if os.path.isdir(path):
            raise RuntimeError(f"无法清空 {path}：里面的文件可能正被 Excel 打开，请关闭后重试")


def write_record(task, page, rows, batches, parts, path):
    lab = TRAITS[task["trait"]]
    append_csv(os.path.join(CONFIG["DOWNLOAD_DIR"], RECORD_NAME),
               ["下载时间", "性状", lab["farm_label"], f"{lab['date_label']}开始日期", f"{lab['date_label']}结束日期",
                "国标编码个数", "编码批数", "文件段数", "页面显示条数（参考）", "文件数据行数", "保存位置"],
               [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), task["trait"], task["farm"] or "不限",
                task["start"] or "不限", task["end"] or "不限", len(task["codes"]) or "不限", batches, parts,
                "" if page is None else page, "" if rows is None else rows,
                rel_to(path, CONFIG["DOWNLOAD_DIR"]) if path else "（无数据，未导出）"])


def download_task(driver, task):
    """下载一个任务，返回 {"path", "total", "rows", "parts"}；没有数据返回 None。"""
    task = norm_task(task)
    t0 = time.time()
    today = datetime.date.today()
    if task["end"] and _date(task["end"]) > today:
        if _date(task["start"]) > today:
            log(f"{TRAITS[task['trait']]['date_label']}日期 {task['start']} 还没到，跳过")
            return None
        log(f"结束日期 {task['end']} 晚于今天，按今天 {today} 查询")
        task["end"] = today.isoformat()
    log(f"---- 开始：{task_name(task)} ----")
    log(f"条件：{describe_conditions(task)}")
    a, b = _date(task["start"]), _date(task["end"])
    base = period_folder(CONFIG["DOWNLOAD_DIR"], a, b) if a else period_folder(CONFIG["DOWNLOAD_DIR"], today, today)
    stem = make_stem(task)
    seg_dir = os.path.join(base, stem + "_分段")
    work = os.path.join(base, f".{stem}_临时")
    _rmtree(work)
    os.makedirs(work, exist_ok=True)

    batches = code_batches(task["codes"])
    if len(batches) > 1:
        log(f"国标编码 {len(task['codes'])} 个，超过 {CONFIG['CODE_BATCH']} 个：分 {len(batches)} 批导出，最后合并成一个文件")
    pieces, ctx = [], {}
    try:
        for k, batch in enumerate(batches, 1):
            if len(batches) > 1:
                log(f"== 第 {k}/{len(batches)} 批：{len(batch)} 个编码 ==")
                core.progress(k - 1, len(batches), "批")
            cond = {"trait": task["trait"], "farm": task["farm"], "start": task["start"], "end": task["end"],
                    "codes": batch or [], "batch": k}
            _collect(driver, cond, work, pieces, ctx)
        if len(batches) > 1:
            core.progress(len(batches), len(batches), "批")

        rows = sum(p["rows"] for p in pieces) if all(p["rows"] is not None for p in pieces) else None
        pages = [p["page"] for p in pieces]
        page = sum(pages) if pieces and all(x is not None for x in pages) else None
        for old in _same_task_files(base, stem):           # 同一任务之前下载的文件已过时
            try:
                os.remove(old)
                log(f"已删除同一任务之前下载的文件：{os.path.basename(old)}")
            except OSError:
                pass
        _rmtree(seg_dir)
        if not pieces:
            write_record(task, 0, 0, len(batches), 0, "")
            log("没有数据，未生成文件")
            return None

        if len(pieces) == 1:
            p = pieces[0]
            final = move_replace(p["path"], base, f"{stem}_{_count_text(rows)}", os.path.splitext(p["path"])[1])
        elif len(batches) > 1 and all(p["path"].lower().endswith(".xlsx") for p in pieces):
            log(f"合并 {len(pieces)} 个文件 ...")
            tmp_out = os.path.join(work, "合并.xlsx")
            merged = merge_xlsx([p["path"] for p in pieces], tmp_out)
            if rows is not None and merged != rows:
                log(f"警告：合并后 {merged} 条，各文件合计 {rows} 条，请核对")
            rows = merged
            final = move_replace(tmp_out, base, f"{stem}_{_count_text(rows)}", ".xlsx")
        else:
            os.makedirs(seg_dir, exist_ok=True)
            with open(os.path.join(seg_dir, "分段明细.csv"), "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(["分段", "编码批次", "日期范围", "页面显示条数（参考）", "文件数据行数", "文件"])
                for k, p in enumerate(pieces, 1):
                    rng = f"{p['start']}~{p['end']}" if p["start"] else NO_DATE_TEXT
                    name = sanitize(f"{stem}_{'第' + str(p['batch']) + '批_' if len(batches) > 1 else ''}"
                                    f"{rng}_{_count_text(p['rows'])}")
                    dst = move_replace(p["path"], seg_dir, name, os.path.splitext(p["path"])[1])
                    w.writerow([k, p["batch"], rng, "" if p["page"] is None else p["page"],
                                "" if p["rows"] is None else p["rows"], os.path.basename(dst)])
                w.writerow(["合计", "", date_text(task), "" if page is None else page,
                            "" if rows is None else rows, ""])
            final = seg_dir

        und = ctx.get("undated_total")
        if und and rows is not None and rows < und:
            log(f"警告：不限日期时页面共 {und} 条，按日期分段导出了 {rows} 条，"
                f"另有 {und - rows} 条的{TRAITS[task['trait']]['date_label']}日期为空或早于 {CONFIG['SPLIT_START']}，未包含在文件里")
        write_record(task, page, rows, len(batches), len(pieces), final)
        size = f"{os.path.getsize(final) / 1024:.0f} KB，" if os.path.isfile(final) else ""
        log(f"[完成] {task_name(task)} -> {final}（{size}{rows if rows is not None else page} 条，"
            f"用时 {fmt_secs(time.time() - t0)}）")
        return {"path": final, "total": rows if rows is not None else page, "rows": rows,
                "parts": len(pieces) if final == seg_dir else 1}
    finally:
        shutil.rmtree(work, ignore_errors=True)


# =====================================================================
# 会话 / 主流程
# =====================================================================
def _goto(driver):
    _MAIN["handle"] = None
    goto_page(driver)


def open_session(username, password):
    return core.start_session(CONFIG["DOWNLOAD_DIR"], username, password, _goto)


def reload_page(driver):
    """不重新登录，重新打开遗传性状查询页面（清掉残留弹窗和多余窗口）。"""
    core.close_other_windows(driver, _MAIN.get("handle"))
    open_url(driver, CONFIG["GSM_URL"])
    if not wait_until(lambda: on_page(driver), 20):
        core.back_to_home(driver)
        goto_page(driver)
    net_hook(driver)
    _MAIN["handle"] = _current_handle(driver)


ADAPTER = core.Adapter(goto=_goto, run=download_task, reload=reload_page, name=task_name,
                       download_dir=lambda: CONFIG["DOWNLOAD_DIR"], title=TITLE)


def make_task(farm="", start="", end="", codes=None, trait=DEFAULT_TRAIT):
    return norm_task({"trait": trait, "farm": farm, "start": start, "end": end, "codes": list(codes or [])})


def run(tasks, username=None, password=None, status_fn=None):
    tasks = [norm_task(t) for t in tasks]
    return core.run_tasks(tasks, ADAPTER, username, password, status_fn, what="任务")


def parse_tasks(spec, codes=None):
    """'南沙母猪场16线,2026-08-01,2026-08-31;,2026-09-01,2026-09-30' → 任务列表（场留空 = 不限，日期留空 = 不限）"""
    out = []
    for item in spec.split(";"):
        if not item.strip():
            continue
        p = [x.strip() for x in re.split(r"[,，]", item)]
        p += [""] * (3 - len(p))
        out.append(make_task(p[0], p[1], p[2], codes))
    return out


def main():
    if "--headless" in sys.argv:
        core.HEADLESS = True
    codes = []
    if "--codes" in sys.argv:
        i = sys.argv.index("--codes")
        if i + 1 < len(sys.argv):
            codes = codes_from_file(sys.argv[i + 1])
    tasks = CONFIG["TASKS"]
    if "--tasks" in sys.argv:
        i = sys.argv.index("--tasks")
        if i + 1 < len(sys.argv):
            tasks = parse_tasks(sys.argv[i + 1], codes)
    username, password = core.load_credentials()
    if tasks:
        run(tasks, username, password)


if __name__ == "__main__":
    main()
