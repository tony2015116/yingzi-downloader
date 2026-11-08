# -*- coding: utf-8 -*-
"""
猪只转移 —— 逐条导出转移单详情
================================================================
首页「养殖生产」→「养殖公共」→「生产管理」→「猪只转移」→ 每个任务（场 × 时间段）：
选场 → 重置 → 填起止日期 → 搜索 → 逐行（含翻页）点「详情」→ 抽屉里「导出」→ 按该行信息命名，
并生成「转移记录清单.csv」和「汇总_转移记录与档案明细.xlsx」。已导出的单据重跑时自动跳过（断点续传）。

    python task_transfer.py              # 按 CONFIG["FILTER_GROUPS"] 下载
    python task_transfer.py --headless
    python task_transfer.py --options    # 只读取各下拉框的可选项并打印
"""

import os
import re
import sys
import csv
import time
import datetime
from collections import ChainMap

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    InvalidSessionIdException,
    WebDriverException,
)

import yz_core as core
from yz_core import (StopRequested, log, wait_until, fmt_secs, sanitize, safe_click, _js, _raise_if_dead,
                     _handles, _current_handle, _n as _norm, _digits, JS_VIS, JS_CLICK_BY_TEXT, JS_SET_VALUE,
                     JS_VISIBLE_TEXT, JS_ANY_VISIBLE, SPIN_CSS, net_hook, net_reset, net_state, spinning,
                     wait_loading_done, switch_to_new_window, confirm_dialog_if_any, save_debug, open_url,
                     dismiss_guide, query_total, wait_query_done, watch_dirs, list_files, recent_partials,
                     file_stable, move_unique, check_stop)

TITLE = "猪只转移"
MANUAL_HOOK = None     # callable(message) -> bool：手动设置参数模式下，等你在浏览器里设置好；返回 False = 取消

CONFIG = ChainMap({
    "DOWNLOAD_DIR": r"D:\test\downloads\猪只转移",
    "SEARCH_TIMEOUT": 180,      # 点「搜索」后最多等多久
    "SEARCH_IDLE_GRACE": 10,
    "DETAIL_TIMEOUT": 60,       # 点「详情」后等弹窗及内容加载的最长时间
    "DOWNLOAD_TIMEOUT": 600,
    "DOWNLOAD_IDLE_RETRY": 40,  # 点「导出」后这么久没有任何动静，就再点一次
    # 手动设置参数：进入页面后暂停，由你在浏览器里选好场和其它参数；程序只填时间、搜索、逐条导出（会显示浏览器）
    "MANUAL_PARAMS": False,
    "REFRESH_RECENT_MONTHS": 1,     # 最近几个整月即使已下完也重新检查（可能有补录的单据）；0 = 都跳过
    "RESET_BETWEEN_TASKS": True,
    "MAX_ROWS": 5000,           # 单个任务最多下载多少条详情
    "FILTER_GROUPS": [
        {"name": "信宜场_202610", "filters": {"所在场": "信宜场",
                                              "起始转移时间": "2026-10-01", "截止转移时间": "2026-10-08"}},
    ],
}, core.CONFIG)

FARM_FIELD = "所在场"
SELECT_FIELDS = ["转移类型", "转移方式", "转入场", "转出场", "猪只种类"]
PARAM_FIELDS = [FARM_FIELD] + SELECT_FIELDS
TIME_FIELDS = ["起始转移时间", "截止转移时间"]
TEXT_FIELDS = ["单据编号", "最少转移头数", "最多转移头数"]
ALL_VALUES = ("", "全部", "不限", "所有")
KEEP_FARM = ("", "全部", "当前场", "当前", "默认")
FIELD_ALIASES = {"猪只类型": "猪只种类", "场": FARM_FIELD, "猪场": FARM_FIELD}


def normalize_filters(filters):
    return {FIELD_ALIASES.get(k, k): v for k, v in (filters or {}).items()}


def make_group(farm, start, end):
    return {"name": f"{farm}_{period_tag(start, end)}",
            "filters": {FARM_FIELD: farm, "起始转移时间": start, "截止转移时间": end}}


# =====================================================================
# 窗口 / iframe
# =====================================================================
_WIN = {"page": None, "frame": []}

JS_FIND_TEXT = JS_VIS + """
const texts = arguments[0];
const all = Array.from(document.querySelectorAll('a, span, li, div, p, button, label, b, em, h3, h4'));
for (const t of texts) {
  const hits = all.filter(e => norm(e.textContent) === t && vis(e));
  const inner = hits.filter(e => !hits.some(o => o !== e && e.contains(o)));
  if (inner.length) return inner[0];
}
return null;
"""


def _enter_path(driver, path):
    """切到顶层页面，再按下标依次进入 iframe。"""
    try:
        driver.switch_to.default_content()
        for i in path:
            frames = driver.find_elements(By.TAG_NAME, "iframe")
            if i >= len(frames):
                return False
            driver.switch_to.frame(frames[i])
        return True
    except WebDriverException as e:
        _raise_if_dead(e)
        return False


def _frame_paths(driver):
    paths = [[]]
    try:
        driver.switch_to.default_content()
        n = len(driver.find_elements(By.TAG_NAME, "iframe"))
        for i in range(n):
            paths.append([i])
            if _enter_path(driver, [i]):
                m = len(driver.find_elements(By.TAG_NAME, "iframe"))
                paths += [[i, j] for j in range(m)]
        driver.switch_to.default_content()
    except WebDriverException as e:
        _raise_if_dead(e)
    return paths


def _find_anywhere(driver, texts):
    """在顶层和各 iframe 中找文字元素。找到时停留在该元素所在上下文。"""
    for p in _frame_paths(driver):
        if not _enter_path(driver, p):
            continue
        el = _js(driver, JS_FIND_TEXT, texts)
        if el is not None:
            return el
    driver.switch_to.default_content()
    return None


def click_anywhere(driver, texts, timeout):
    el = wait_until(lambda: _find_anywhere(driver, texts), timeout, poll=0.5)
    if el is None:
        return False
    safe_click(driver, el)
    try:
        driver.switch_to.default_content()
    except WebDriverException as e:
        _raise_if_dead(e)
    return True


def confirm_msgbox(driver):
    """只处理确认类小弹框（避免误点详情弹窗里的「确定」把详情关掉）。"""
    return confirm_dialog_if_any(driver, ".el-message-box__wrapper, .ant-modal-confirm, .ant-popover")


# =====================================================================
# 导航：养殖生产 → 养殖公共 → 生产管理 → 猪只转移
# =====================================================================
JS_HAS_FORM = """
const t = (document.body && document.body.innerText) || '';
return t.includes('转移类型') && /猪只种类|猪只类型/.test(t) && /搜\\s*索|查\\s*询/.test(t);
"""


def _scan_form(driver):
    for p in _frame_paths(driver):
        if _enter_path(driver, p) and _js(driver, JS_HAS_FORM, default=False):
            return ("ok", p)
    driver.switch_to.default_content()
    return None


def locate_form(driver, timeout):
    r = wait_until(lambda: _scan_form(driver), timeout, poll=0.6)
    return r[1] if r else None


def goto_transfer_page(driver, max_attempts=3):
    for attempt in range(1, max_attempts + 1):
        try:
            driver.switch_to.default_content()
            path = locate_form(driver, 1)
            if path is None:
                driver.switch_to.default_content()
                if _find_anywhere(driver, ["养殖公共"]) is None:
                    log("进入「养殖生产」模块 ...")
                    before = set(_handles(driver))
                    if not click_anywhere(driver, ["养殖生产"], CONFIG["TIMEOUT"]):
                        raise RuntimeError("首页未找到「养殖生产」")
                    switch_to_new_window(driver, before, wait=3)
                    wait_loading_done(driver, 60)

                log("点击顶部「养殖公共」...")
                before = set(_handles(driver))
                if not click_anywhere(driver, ["养殖公共"], 40):
                    raise RuntimeError("未找到顶部「养殖公共」")
                switch_to_new_window(driver, before, wait=1.5)
                wait_loading_done(driver, 30)

                if not wait_until(lambda: _find_anywhere(driver, ["猪只转移"]), 4, poll=0.5):
                    log("展开左侧「生产管理」...")
                    if not click_anywhere(driver, ["生产管理"], 20):
                        raise RuntimeError("左侧菜单未找到「生产管理」")
                    if not wait_until(lambda: _find_anywhere(driver, ["猪只转移"]), 6, poll=0.5):
                        click_anywhere(driver, ["生产管理"], 5)   # 可能把已展开的菜单收起了，再点一次
                log("点击「猪只转移」...")
                before = set(_handles(driver))
                if not click_anywhere(driver, ["猪只转移"], 20):
                    raise RuntimeError("「生产管理」下未找到「猪只转移」")
                switch_to_new_window(driver, before, wait=1.5)
                path = locate_form(driver, CONFIG["TIMEOUT"] + 15)
            if path is not None:
                _WIN["page"] = _current_handle(driver)
                _WIN["frame"] = path
                net_hook(driver)
                wait_loading_done(driver, 20)
                log("已进入「猪只转移」页面" + (f"（内容在第 {path} 层 iframe）" if path else ""))
                return True
            log("未检测到「转移类型」等筛选字段，准备重试 ...")
        except (StopRequested, InvalidSessionIdException):
            raise
        except Exception as e:
            _raise_if_dead(e)
            log(f"导航第 {attempt} 次尝试异常: {type(e).__name__}: {str(e)[:100]}")
        save_debug(driver, f"导航失败_{attempt}")
        time.sleep(2)
        if core.SESSION.get("home_url") and attempt < max_attempts:
            open_url(driver, core.SESSION["home_url"])
            wait_until(lambda: core._at_home(driver), 20)
    raise RuntimeError("多次尝试后仍未能进入「猪只转移」页面，请检查菜单名称是否有变化。")


def ensure_page(driver):
    """确认当前位于猪只转移页的内容上下文；不在则切回。"""
    ph = _WIN.get("page")
    if ph and _current_handle(driver) != ph:
        if ph not in _handles(driver):
            raise RuntimeError("猪只转移页面窗口已被关闭")
        driver.switch_to.window(ph)
    if _enter_path(driver, _WIN.get("frame") or []) and _js(driver, JS_HAS_FORM, default=False):
        net_hook(driver)
        return
    path = locate_form(driver, 10)
    if path is None:
        raise RuntimeError("找不到猪只转移页面的筛选区")
    _WIN["frame"] = path
    net_hook(driver)


# =====================================================================
# 筛选字段
# =====================================================================
SELECT_CSS = (".el-select, .el-cascader, .ant-select, .ant-cascader-picker, .ant-cascader, "
              ".vue-treeselect, .el-input, input")
OPT_CSS = (".el-select-dropdown__item, .ant-select-item-option, .el-cascader-node, .ant-cascader-menu-item, "
           ".vue-treeselect__option, .el-tree-node__content, li[role='option']")
POPUP_CSS = (".el-select-dropdown, .el-picker-panel, .el-cascader__dropdown, .el-cascader-panel, "
             ".ant-select-dropdown:not(.ant-select-dropdown-hidden), "
             ".ant-picker-dropdown:not(.ant-picker-dropdown-hidden), .ant-cascader-menus, .vue-treeselect__menu")
PANEL_CSS = ".el-picker-panel, .ant-picker-dropdown:not(.ant-picker-dropdown-hidden)"

JS_FIELD_CTL = JS_VIS + """
const label = arguments[0], css = arguments[1];
const clean = s => (s || '').replace(/[\\s:：*＊]/g, '');
const bad = e => e.closest('table, .el-table, .ant-table, .el-dialog__wrapper, .ant-modal-wrap, .el-drawer__wrapper, '
                         + '.el-picker-panel, .el-select-dropdown, .ant-select-dropdown, .ant-picker-dropdown');
const pick = root => { for (const c of root.querySelectorAll(css)) if (vis(c) && !bad(c)) return c; return null; };
for (const lb of document.querySelectorAll('.el-form-item__label, .ant-form-item-label, label')) {
  if (bad(lb) || !vis(lb) || clean(lb.textContent) !== label) continue;
  const item = lb.closest('.el-form-item, .ant-form-item');
  if (item) { const c = pick(item); if (c) return c; }
}
for (const ph of document.querySelectorAll('.ant-select-selection-placeholder, .ant-select-selection__placeholder, '
                                           + '.el-select__placeholder, .vue-treeselect__placeholder')) {
  if (bad(ph) || !vis(ph) || clean(ph.textContent) !== label) continue;
  const c = ph.closest('.ant-select, .el-select, .ant-cascader-picker, .ant-cascader, .vue-treeselect');
  if (c) return c;
}
for (const i of document.querySelectorAll('input[placeholder]')) {
  if (bad(i) || !vis(i)) continue;
  const p = clean(i.getAttribute('placeholder')).replace(/^请(选择|输入)/, '');
  if (p === label) return i.matches(css) ? i : (i.closest(css.replace(/,\\s*input$/, '')) || i);
}
const own = e => Array.from(e.childNodes).filter(n => n.nodeType === 3).map(n => n.textContent).join('');
for (const e of document.querySelectorAll('span, div, label, p, b')) {
  if (bad(e) || !vis(e) || clean(own(e)) !== label) continue;
  let p = e;
  for (let k = 0; k < 5 && p.parentElement; k++) {
    p = p.parentElement;
    const cs = Array.from(p.querySelectorAll(css))
      .filter(c => vis(c) && !bad(c) && (e.compareDocumentPosition(c) & 4));
    if (cs.length) return cs[0];
  }
}
return null;
"""

JS_CTL_TEXT = """
const c = arguments[0], parts = [];
if (c.tagName === 'INPUT' && c.value) parts.push(c.value);
c.querySelectorAll('input').forEach(i => { if (i.value) parts.push(i.value); });
c.querySelectorAll('.el-select__tags-text, .el-tag, .ant-select-selection-item, .vue-treeselect__single-value, '
                   + '.vue-treeselect__multi-value-label').forEach(t => { const s = (t.textContent || '').trim(); if (s) parts.push(s); });
return Array.from(new Set(parts)).join('、');
"""

JS_PICK_OPTION = JS_VIS + """
const want = norm(arguments[1]);
const opts = Array.from(document.querySelectorAll(arguments[0])).filter(e => vis(e)
  && !e.classList.contains('is-disabled') && !e.classList.contains('ant-select-item-option-disabled'));
let hit = opts.find(o => norm(o.textContent) === want);
if (hit) return hit;
const part = opts.filter(o => norm(o.textContent).includes(want));
return part.length === 1 ? part[0] : null;
"""

JS_OPTION_TEXTS = JS_VIS + """
return Array.from(document.querySelectorAll(arguments[0])).filter(vis)
  .map(e => (e.textContent || '').replace(/\\s+/g, ' ').trim()).filter(Boolean);
"""

JS_SCROLL_DROPDOWN = JS_VIS + """
let moved = false;
for (const h of document.querySelectorAll('.el-select-dropdown .el-scrollbar__wrap, .rc-virtual-list-holder, '
                                          + '.ant-select-dropdown .rc-virtual-list-holder, .vue-treeselect__menu')) {
  if (!vis(h)) continue;
  const t = h.scrollTop; h.scrollTop = t + Math.max(h.clientHeight - 20, 40);
  h.dispatchEvent(new Event('scroll', {bubbles: true}));
  if (h.scrollTop !== t) moved = true;
}
return moved;
"""

JS_CLEAR_CTL = """
const c = arguments[0]; let n = 0;
c.querySelectorAll('.el-tag__close, .ant-select-selection-item-remove, .vue-treeselect__value-remove').forEach(x => { x.click(); n++; });
for (const x of c.querySelectorAll('.el-icon-circle-close, .el-input__clear, .ant-select-clear, .vue-treeselect__x-container')) {
  ['mousedown', 'mouseup', 'click'].forEach(t => x.dispatchEvent(new MouseEvent(t, {bubbles: true})));
  n++; break;
}
return n;
"""

JS_BODY_MOUSE = """
for (const t of ['mousedown', 'mouseup', 'click']) document.body.dispatchEvent(new MouseEvent(t, {bubbles: true}));
"""

JS_FARM_CTL = JS_VIS + """
const fs = Array.from(document.querySelectorAll('.fpf-components-org-farm-selector-index-wrapper, '
                                               + '[class*="org-farm-selector"][class*="wrapper"]')).find(vis);
if (fs) return fs;
const SEL = '.ant-select, .el-select';
const out = s => !s.closest('form, .ant-form, .el-form, table, .ant-table, .el-table, .ant-modal-wrap, .el-dialog__wrapper');
for (const b of document.querySelectorAll('button, a, .ant-btn, .el-button')) {
  if (!vis(b) || !['转入', '转出'].includes(norm(b.textContent))) continue;
  let p = b;
  for (let k = 0; k < 6 && p.parentElement; k++) {
    p = p.parentElement;
    const s = Array.from(p.querySelectorAll(SEL)).filter(x => vis(x) && out(x));
    if (s.length) return s[0];
  }
}
let best = null, bt = 1e9;
for (const s of document.querySelectorAll(SEL)) {
  if (!vis(s) || !out(s)) continue;
  const t = s.getBoundingClientRect().top;
  if (t < bt) { bt = t; best = s; }
}
return best;
"""

_FIELD_CACHE = {}   # 字段名 → 控件。选了值以后占位文字会消失，靠缓存继续找到同一个框


def _field_ctl(driver, label, css, timeout=8):
    def alive(el):
        try:
            return el.is_displayed()
        except WebDriverException:
            return False

    el = _js(driver, JS_FIELD_CTL, label, css)
    if el is None:
        cached = _FIELD_CACHE.get((label, css))
        if cached is not None and alive(cached):
            return cached
        el = wait_until(lambda: _js(driver, JS_FIELD_CTL, label, css), timeout, poll=0.3)
    if el is not None:
        _FIELD_CACHE[(label, css)] = el
    return el


def _ctl_text(driver, ctl):
    return _js(driver, JS_CTL_TEXT, ctl, default="") or ""


def _pick(driver, value):
    return _js(driver, JS_PICK_OPTION, OPT_CSS, value)


def _options_open(driver):
    return bool(_js(driver, JS_ANY_VISIBLE, OPT_CSS, default=False))


def _popup_open(driver):
    return bool(_js(driver, JS_ANY_VISIBLE, POPUP_CSS, default=False))


def close_popups(driver):
    if not _popup_open(driver):
        return
    try:
        driver.switch_to.active_element.send_keys(Keys.ESCAPE)
    except Exception:
        pass
    if wait_until(lambda: not _popup_open(driver), 0.6, poll=0.15):
        return
    _js(driver, JS_BODY_MOUSE)
    wait_until(lambda: not _popup_open(driver), 1.0, poll=0.15)


def _open_dropdown(driver, ctl):
    safe_click(driver, ctl)
    if wait_until(lambda: _options_open(driver), 2.5, poll=0.2):
        return True
    for inp in ctl.find_elements(By.CSS_SELECTOR, "input"):
        try:
            if inp.is_displayed():
                inp.click()
                break
        except WebDriverException:
            pass
    return bool(wait_until(lambda: _options_open(driver), 2.5, poll=0.2))


def _type_filter(driver, ctl, text):
    cands = ctl.find_elements(By.CSS_SELECTOR, "input")
    if ctl.tag_name.lower() == "input":
        cands.insert(0, ctl)
    cands += driver.find_elements(By.CSS_SELECTOR, ".el-select-dropdown input, .ant-select-dropdown input, "
                                                   ".el-popover input, .vue-treeselect__menu input")
    for i in cands:
        try:
            if i.is_displayed() and not i.get_attribute("readonly"):
                i.send_keys(Keys.CONTROL, "a")
                i.send_keys(text)
                return True
        except WebDriverException:
            continue
    return False


def _select_all(driver, ctl, label):
    shown = _ctl_text(driver, ctl)
    if not shown or "全部" in shown:
        log(f"  {label} = 全部")
        return
    if _open_dropdown(driver, ctl):
        el = _pick(driver, "全部")
        if el is not None:
            safe_click(driver, el)
            close_popups(driver)
            log(f"  {label} = 全部")
            return
    close_popups(driver)
    try:
        ActionChains(driver).move_to_element(ctl).perform()
    except Exception:
        pass
    _js(driver, JS_CLEAR_CTL, ctl)
    close_popups(driver)
    if wait_until(lambda: not _ctl_text(driver, ctl), 1.5, poll=0.2):
        log(f"  {label} 已清空（= 全部）")
    else:
        log(f"  警告：{label} 未能清空，当前为「{_ctl_text(driver, ctl)}」")


def set_dropdown(driver, label, value):
    value = str(value or "").strip()
    ctl = _field_ctl(driver, label, SELECT_CSS)
    if ctl is None:
        raise RuntimeError(f"未找到下拉字段「{label}」")
    if value in ALL_VALUES:
        _select_all(driver, ctl, label)
        return
    _choose_option(driver, ctl, label, value)


def _choose_option(driver, ctl, label, value):
    if _norm(_ctl_text(driver, ctl)) == _norm(value.split("/")[-1]):
        log(f"  {label} = {value}（已是该值）")
        return
    parts = [p.strip() for p in value.split("/") if p.strip()] or [value]
    _open_dropdown(driver, ctl)
    for part in parts:
        el = wait_until(lambda: _pick(driver, part), 2.5, poll=0.2)
        if el is None and _type_filter(driver, ctl, part):
            el = wait_until(lambda: _pick(driver, part), 4, poll=0.25)
        if el is None:
            close_popups(driver)
            save_debug(driver, f"选项不存在_{label}")
            raise RuntimeError(f"下拉字段「{label}」中未找到选项「{part}」")
        safe_click(driver, el)
        time.sleep(0.25)          # 级联选择：等下一列展开
    close_popups(driver)
    shown = _ctl_text(driver, ctl)
    if _norm(parts[-1]) in _norm(shown):
        log(f"  {label} = {value}")
    else:
        log(f"  警告：{label} 当前显示「{shown}」，期望「{value}」，请核对下载结果")


def _farm_ctl(driver, timeout=10):
    return wait_until(lambda: _js(driver, JS_FARM_CTL), timeout, poll=0.4)


# 右上角的场选择是平台自定义组件（输入框 + 弹出层）
JS_FARM_ITEMS = JS_VIS + """
const w = arguments[0], want = arguments[1] ? norm(arguments[1]) : null;
const roots = [w, ...Array.from(document.querySelectorAll('.ant-popover, .ant-dropdown, .ant-select-dropdown, '
                                                        + '[class*="farm-selector"][class*="popup"]')).filter(vis)];
const items = [];
for (const r of roots) for (const e of r.querySelectorAll('*')) {
  if (e.tagName === 'INPUT' || e.tagName === 'svg' || e.closest('svg') || !vis(e)) continue;
  const t = norm(e.textContent);
  if (!t || t.length > 40) continue;
  if (Array.from(e.children).some(c => norm(c.textContent) === t)) continue;
  if (e.closest('input') || e.classList.contains('ant-input-suffix')) continue;
  if (want === null) items.push(e.textContent.trim());
  else if (t === want) return e;
}
return want === null ? Array.from(new Set(items)) : null;
"""


def _is_farm_widget(driver, ctl):
    return bool(_js(driver, "return /farm-selector/.test(arguments[0].className || '');", ctl, default=False))


def _farm_input(ctl):
    ins = ctl.find_elements(By.CSS_SELECTOR, "input")
    return ins[0] if ins else ctl


def _open_farm_widget(driver, ctl):
    safe_click(driver, _farm_input(ctl))
    return wait_until(lambda: len(_js(driver, JS_FARM_ITEMS, ctl, None, default=[]) or []) > 1, 4, poll=0.3)


def _choose_farm_widget(driver, ctl, value):
    name = value.split("/")[-1].strip()
    _open_farm_widget(driver, ctl)
    el = wait_until(lambda: _js(driver, JS_FARM_ITEMS, ctl, name), 2, poll=0.3)
    if el is None:
        boxes = [i for i in driver.find_elements(By.CSS_SELECTOR, ".ant-popover input, .ant-dropdown input, "
                                                                  "[class*='farm-selector'] input")
                 if i.is_displayed() and not i.get_attribute("readonly")]
        for box in boxes[::-1]:
            try:
                box.send_keys(Keys.CONTROL, "a")
                box.send_keys(name)
                break
            except WebDriverException:
                continue
        el = wait_until(lambda: _js(driver, JS_FARM_ITEMS, ctl, name), 5, poll=0.3)
    if el is None:
        close_popups(driver)
        save_debug(driver, f"场选择失败_{name}")
        raise RuntimeError(f"右上角的场选择中未找到「{name}」（场名需与页面上显示的完全一致）")
    safe_click(driver, el)
    if wait_until(lambda: _norm(_ctl_text(driver, ctl)) == _norm(name), 8, poll=0.3):
        log(f"  {FARM_FIELD} = {name}")
    else:
        log(f"  警告：{FARM_FIELD} 当前显示「{_ctl_text(driver, ctl)}」，期望「{name}」")


JS_READ_PARAMS = JS_VIS + """
const order = arguments[0], out = {};
const d = document.querySelector('#eventDate');
const form = (d && d.closest('form')) || document.querySelector('.ant-legacy-form, .ant-form, form');
if (form) {
  const sels = Array.from(form.querySelectorAll('.ant-select')).filter(s => vis(s) && !s.closest('.ant-pagination'));
  sels.forEach((s, i) => {
    const ph = s.querySelector('.ant-select-selection-placeholder');
    const item = s.querySelector('.ant-select-selection-item');
    const v = item ? (item.getAttribute('title') || item.textContent || '').trim() : '';
    const key = ph ? (ph.textContent || '').trim() : order[i];
    if (key) out[key] = v;
  });
  for (const i of form.querySelectorAll('input[placeholder]')) {
    const p = (i.getAttribute('placeholder') || '').trim();
    if (p && i.value && !i.closest('.ant-select')) out[p] = i.value;
  }
}
const farm = Array.from(document.querySelectorAll('[class*="org-farm-selector"] input')).find(vis);
if (farm) out['__farm'] = farm.value;
return out;
"""


def read_current_params(driver):
    return _js(driver, JS_READ_PARAMS, SELECT_FIELDS, default={}) or {}


def describe_params(p):
    items = [f"{FARM_FIELD}={p.get('__farm') or '?'}"]
    items += [f"{k}={p.get(k) or '全部'}" for k in SELECT_FIELDS]
    extra = [f"{k}={v}" for k, v in p.items() if k not in SELECT_FIELDS + TIME_FIELDS + ["__farm"] and v]
    return "，".join(items + extra)


def wait_manual(driver, msg):
    ensure_page(driver)
    log(msg)
    if MANUAL_HOOK is not None:
        if not MANUAL_HOOK(msg):
            raise StopRequested("用户取消")
    else:
        input("设置好后回到这里按回车继续 ...")
    check_stop()
    close_popups(driver)
    ensure_page(driver)
    log(f"当前页面参数：{describe_params(read_current_params(driver))}")


def set_farm(driver, value):
    """先选页面右上角的场。返回当前选中的场名。"""
    value = str(value or "").strip()
    ctl = _farm_ctl(driver)
    if ctl is None:
        if value in KEEP_FARM:
            log("  未找到右上角的场选择框，保持页面当前状态")
            return ""
        save_debug(driver, "未找到场选择框")
        raise RuntimeError("未找到页面右上角的场选择框")
    cur = _ctl_text(driver, ctl)
    if value in KEEP_FARM or _norm(cur) == _norm(value.split("/")[-1]):
        log(f"  {FARM_FIELD} = {cur or '（未选）'}")
        return cur
    net_reset(driver)
    if _is_farm_widget(driver, ctl):
        _choose_farm_widget(driver, ctl, value)
    else:
        _choose_option(driver, ctl, FARM_FIELD, value)
    # 切换场后页面会重新加载数据，等它稳定
    time.sleep(0.5)
    wait_until(lambda: not (net_state(driver) or {}).get("p") and not spinning(driver), 30, poll=0.3)
    wait_until(lambda: _js(driver, JS_HAS_FORM, default=False), 15)
    _FIELD_CACHE.clear()
    ctl = _farm_ctl(driver, 5)
    cur = _ctl_text(driver, ctl) if ctl is not None else ""
    if cur and _norm(cur) != _norm(value.split("/")[-1]):
        raise RuntimeError(f"场选择框当前为「{cur}」，不是「{value}」（切换未生效）")
    return cur or value


def read_options(driver, label, max_scroll=60):
    ctl = _farm_ctl(driver) if label == FARM_FIELD else _field_ctl(driver, label, SELECT_CSS)
    if ctl is None:
        raise RuntimeError(f"未找到下拉字段「{label}」")
    if label == FARM_FIELD and _is_farm_widget(driver, ctl):
        _open_farm_widget(driver, ctl)
        cur = _ctl_text(driver, ctl)
        items = [t for t in (_js(driver, JS_FARM_ITEMS, ctl, None, default=[]) or []) if t != cur]
        items += [cur] if cur else []
        close_popups(driver)
        _js(driver, JS_BODY_MOUSE)
        return items
    if not _open_dropdown(driver, ctl):
        close_popups(driver)
        return []
    seen = []
    for _ in range(max_scroll):
        for t in _js(driver, JS_OPTION_TEXTS, OPT_CSS, default=[]) or []:
            if t not in seen:
                seen.append(t)
        if not _js(driver, JS_SCROLL_DROPDOWN, default=False):
            break
        time.sleep(0.15)
    close_popups(driver)
    return seen


def norm_datetime(v, is_end=False):
    """"2026-10-01" / "2026/10/01 8:00" → "YYYY-MM-DD HH:MM:SS"（只写日期时起始补 00:00:00，截止补 23:59:59）。"""
    nums = [int(x) for x in re.findall(r"\d+", str(v or ""))]
    if len(nums) < 3:
        raise ValueError(f"时间格式不正确：{v}")
    y, mo, d = nums[:3]
    datetime.date(y, mo, d)          # 校验日期本身是否合法（旧版 2026-02-30 这类会到页面上才报错）
    if len(nums) >= 4:
        hh, mm, ss = (nums[3:6] + [0, 0])[:3]
    else:
        hh, mm, ss = (23, 59, 59) if is_end else (0, 0, 0)
    return f"{y:04d}-{mo:02d}-{d:02d} {hh:02d}:{mm:02d}:{ss:02d}"


def _datetime_via_panel(driver, inp, label, value):
    safe_click(driver, inp)
    panel = wait_until(lambda: next((p for p in driver.find_elements(By.CSS_SELECTOR, PANEL_CSS)
                                     if p.is_displayed()), None), 2, poll=0.2)
    if panel is None:
        return
    ins = [i for i in panel.find_elements(By.CSS_SELECTOR, "input") if i.is_displayed()]
    if len(ins) < 2:
        return
    k = 2 if label.startswith("截止") and len(ins) >= 4 else 0
    for el, v in ((ins[k], value[:10]), (ins[k + 1], value[11:])):
        driver.execute_script(JS_SET_VALUE, el, v)
        time.sleep(0.2)
    _js(driver, JS_CLICK_BY_TEXT, ["确定", "确认", "OK", "Ok"], PANEL_CSS, "button, a, span")


def set_datetime(driver, label, value):
    """单独设置一个时间框（起止时间都给出时用 set_date_range 点日历）。"""
    value = norm_datetime(value, is_end=label.startswith("截止"))
    want = _digits(value)
    inp = _field_ctl(driver, label, "input")
    if inp is None:
        raise RuntimeError(f"未找到时间字段「{label}」")

    def state():
        d = _digits(inp.get_attribute("value"))
        if d == want or (len(d) >= 12 and want.startswith(d)):
            return "full"
        if d == want[:8]:
            return "date"
        return None

    in_range = bool(_js(driver, "return !!arguments[0].closest('.ant-picker-range, .ant-calendar-range-picker, "
                                ".el-range-editor, .el-date-editor--datetimerange');", inp, default=False))

    def type_in(text):
        if inp.get_attribute("readonly"):
            return
        try:
            safe_click(driver, inp)
            inp.send_keys(Keys.CONTROL, "a")
            inp.send_keys(Keys.BACKSPACE)
            inp.send_keys(text)
            inp.send_keys(Keys.ENTER)
        except WebDriverException as e:
            _raise_if_dead(e)

    def js_in(text):
        _js(driver, JS_SET_VALUE, inp, text)
        try:
            inp.send_keys(Keys.ENTER)
        except WebDriverException:
            pass

    def press_ok():
        _js(driver, "const b = Array.from(document.querySelectorAll('.ant-picker-ok button, .ant-calendar-ok-btn'))"
                    ".find(x => x.offsetWidth && !x.disabled); if (b) { b.click(); return true; } return false;")

    def ok():
        st = state()
        return st == "full" or (st == "date" and len(_digits(inp.get_attribute("value"))) == 8)

    for cand in (value[:10], value, value[:16]):
        if ok():
            break
        for how in (type_in, js_in):
            if ok():
                break
            how(cand)
            if not wait_until(ok, 1.2, poll=0.15):
                press_ok()
                wait_until(ok, 0.8, poll=0.15)
        if ok():
            break
        if cand == value:
            _datetime_via_panel(driver, inp, label, value)
            if wait_until(ok, 1.2, poll=0.15):
                break
    if not (in_range and label.startswith("起始")):
        close_popups(driver)

    st, cur = state(), inp.get_attribute("value") or ""
    if st == "full":
        log(f"  {label} = {cur}")
    elif st == "date":
        log(f"  {label} = {cur}（该时间框只能选到日期）")
    else:
        save_debug(driver, f"时间设置失败_{label}")
        raise RuntimeError(f"「{label}」未能设为 {value}（当前为「{cur}」）")


def set_text(driver, label, value):
    value = str(value or "").strip()
    inp = _field_ctl(driver, label, "input")
    if inp is None:
        raise RuntimeError(f"未找到输入框「{label}」")
    if (inp.get_attribute("value") or "") == value:
        if value:
            log(f"  {label} = {value}")
        return
    for how in ("keys", "js"):
        try:
            if how == "keys":
                safe_click(driver, inp)
                inp.send_keys(Keys.CONTROL, "a")
                inp.send_keys(Keys.BACKSPACE)
                if value:
                    inp.send_keys(value)
            else:
                _js(driver, JS_SET_VALUE, inp, value)
            _js(driver, "arguments[0].blur();", inp)
        except WebDriverException as e:
            _raise_if_dead(e)
        if wait_until(lambda: _norm(inp.get_attribute("value")) == _norm(value), 1.0, poll=0.15):
            log(f"  {label} = {value or '（空）'}")
            return
    raise RuntimeError(f"「{label}」未能填入「{value}」（当前为「{inp.get_attribute('value')}」）")


# ---------------------------------------------------------------------
# 筛选区展开 + 日期范围框（antd RangePicker）日历点选
# ---------------------------------------------------------------------
JS_EXPAND_FILTERS = JS_VIS + """
const inp = document.querySelector('#eventDate')
  || Array.from(document.querySelectorAll('input[placeholder]')).find(i => /转移时间/.test(i.placeholder));
if (inp && vis(inp)) return 'ok';
const b = ['.yz-search-section .more-btn button', '.more-btn button', '[class*="search"] [class*="more"] button',
           '.more-btn'].map(sel => Array.from(document.querySelectorAll(sel)).find(vis)).find(Boolean);
if (!b) return inp ? 'hidden' : 'none';
b.click();
return 'clicked';
"""

PICKER_DD = ".ant-picker-dropdown:not(.ant-picker-dropdown-hidden)"

JS_PICKER_PANELS = JS_VIS + """
const dd = Array.from(document.querySelectorAll(arguments[0])).find(vis);
if (!dd) return null;
const out = [];
for (const p of dd.querySelectorAll('.ant-picker-date-panel, .ant-picker-panel')) {
  if (!vis(p)) continue;
  const c = p.querySelector('td.ant-picker-cell-in-view[title]');
  if (!c) continue;
  const m = c.getAttribute('title').match(/(\\d{4})\\D(\\d{1,2})/);
  if (m) { const ym = [parseInt(m[1]), parseInt(m[2])];
           if (!out.some(x => x[0] === ym[0] && x[1] === ym[1])) out.push(ym); }
}
return out;
"""

JS_PICKER_CELL = JS_VIS + """
const dd = Array.from(document.querySelectorAll(arguments[0])).find(vis);
if (!dd) return null;
for (const c of dd.querySelectorAll('td.ant-picker-cell-in-view')) {
  if (c.getAttribute('title') !== arguments[1] || !vis(c)) continue;
  return {el: c.querySelector('.ant-picker-cell-inner') || c, disabled: c.classList.contains('ant-picker-cell-disabled')};
}
return null;
"""

JS_PICKER_NAV = JS_VIS + """
const dd = Array.from(document.querySelectorAll(arguments[0])).find(vis);
if (!dd) return false;
const b = Array.from(dd.querySelectorAll('.ant-picker-header-' + arguments[1] + '-btn')).find(vis);
if (!b) return false;
b.click();
return true;
"""


def expand_filters(driver):
    """筛选区默认收起时，第二行（转出场、转移时间…）是隐藏的，点「展开」按钮。"""
    r = _js(driver, JS_EXPAND_FILTERS, default="")
    if r == "clicked":
        if wait_until(lambda: _js(driver, JS_EXPAND_FILTERS, default="") == "ok", 3, poll=0.2):
            log("已展开筛选区")
        else:
            _js(driver, JS_EXPAND_FILTERS)          # 可能点成了「收起」，再点一次
            wait_until(lambda: _js(driver, JS_EXPAND_FILTERS, default="") == "ok", 3, poll=0.2)
    return _js(driver, JS_EXPAND_FILTERS, default="") == "ok"


def _pick_day(driver, target, label, clamp_today=False):
    """在已打开的日历里翻到 target（YYYY-MM-DD）所在月份并点击该日。返回实际点选的日期。"""
    tgt = datetime.date.fromisoformat(target)
    for _ in range(80):
        check_stop()
        cell = _js(driver, JS_PICKER_CELL, PICKER_DD, tgt.isoformat())
        if cell:
            if cell.get("disabled"):
                today = datetime.date.today()
                if clamp_today and tgt > today:
                    log(f"  {label} {tgt} 晚于今天且不可选，改为今天 {today}")
                    tgt = today
                    continue
                raise RuntimeError(f"「{label}」日期 {tgt} 在日历里不可选")
            el = cell["el"]
            try:
                el.click()
            except WebDriverException as e:
                _raise_if_dead(e)
                _js(driver, "arguments[0].click();", el)
            return tgt.isoformat()
        panels = _js(driver, JS_PICKER_PANELS, PICKER_DD)
        if not panels:
            raise RuntimeError("日历面板没有打开")
        lo, hi = min(panels), max(panels)
        t = [tgt.year, tgt.month]
        if t < lo:
            diff = (lo[0] - t[0]) * 12 + lo[1] - t[1]
            way = "super-prev" if diff >= 12 else "prev"
        else:
            diff = (t[0] - hi[0]) * 12 + t[1] - hi[1]
            way = "super-next" if diff >= 12 else "next"
        if not _js(driver, JS_PICKER_NAV, PICKER_DD, way):
            if not _js(driver, JS_PICKER_NAV, PICKER_DD, way.replace("super-", "")):
                raise RuntimeError(f"日历上找不到翻页按钮（{way}）")
        before = panels
        wait_until(lambda: _js(driver, JS_PICKER_PANELS, PICKER_DD) != before, 1.5, poll=0.1)
    raise RuntimeError(f"日历翻页次数过多，未能翻到 {target[:7]}")


def set_date_range(driver, start, end):
    """起止转移时间（同一个只读范围框）：打开日历 → 点起始日 → 点截止日。"""
    s, e = norm_datetime(start)[:10], norm_datetime(end, True)[:10]
    expand_filters(driver)
    inp_s = _field_ctl(driver, "起始转移时间", "input")
    inp_e = _field_ctl(driver, "截止转移时间", "input", timeout=2)
    if inp_s is None:
        save_debug(driver, "未找到时间框")
        raise RuntimeError("未找到时间字段「起始转移时间」（筛选区可能没有展开）")

    def vals():
        a = (inp_s.get_attribute("value") or "")[:10]
        b = (inp_e.get_attribute("value") or "")[:10] if inp_e is not None else ""
        return a, b

    if vals() == (s, e):
        log(f"  转移时间 = {s} ~ {e}（已是该值）")
        return
    for attempt in range(2):
        safe_click(driver, inp_s)
        if not wait_until(lambda: _js(driver, JS_PICKER_PANELS, PICKER_DD), 3, poll=0.2):
            continue
        try:
            got_s = _pick_day(driver, s, "起始转移时间")
            time.sleep(0.2)
            got_e = _pick_day(driver, e, "截止转移时间", clamp_today=True)
        except InvalidSessionIdException:
            raise
        except Exception as ex:
            _raise_if_dead(ex)
            log(f"  日历点选出错：{str(ex)[:80]}，重试")
            close_popups(driver)
            continue
        wait_until(lambda: not _js(driver, JS_PICKER_PANELS, PICKER_DD), 2, poll=0.15)
        expect = (got_s, got_e) if inp_e is not None else (got_s, "")
        if wait_until(lambda: vals() == expect, 2, poll=0.15):
            log(f"  转移时间 = {got_s} ~ {got_e}")
            return
        log(f"  警告：时间框当前为 {vals()}，期望 {got_s} ~ {got_e}，重试")
        close_popups(driver)
    save_debug(driver, "时间设置失败")
    raise RuntimeError(f"未能把转移时间设为 {s} ~ {e}（当前 {vals()}）")


def apply_filter(driver, field, value):
    if field in TEXT_FIELDS:
        set_text(driver, field, value)
    elif field in TIME_FIELDS or "时间" in field or "日期" in field:
        set_datetime(driver, field, value)
    else:
        set_dropdown(driver, field, value)


# =====================================================================
# 搜索 / 表格
# =====================================================================
DIALOG_CSS = (".el-dialog__wrapper, .el-drawer__wrapper, .ant-modal-wrap, .ant-drawer-open, "
              ".el-message-box__wrapper, [role='dialog']")

JS_FIND_BTN = JS_VIS + """
const root = arguments[0] || document, exact = arguments[1], part = arguments[2], skipTable = arguments[3];
const els = Array.from(root.querySelectorAll('button, a, .el-button, [role=button], span'))
  .filter(e => vis(e) && !(skipTable && e.closest('tbody'))
          && !(arguments[0] === null && e.closest('.el-dialog__wrapper, .ant-modal-wrap, .el-drawer__wrapper')));
const cand = els.filter(e => !els.some(o => o !== e && o.contains(e) && norm(o.textContent) === norm(e.textContent)));
const ok = e => !(e.disabled || e.classList.contains('is-disabled') || e.classList.contains('ant-btn-disabled')
                  || e.getAttribute('aria-disabled') === 'true' || (e.closest('button') || {}).disabled);
for (const t of exact) { const h = cand.find(e => norm(e.textContent) === t); if (h) return [h, ok(h)]; }
for (const t of part) {
  const h = cand.find(e => { const s = norm(e.textContent); return s.includes(t) && s.length <= t.length + 6; });
  if (h) return [h, ok(h)];
}
return null;
"""

JS_T_STATE = JS_VIS + """
const q = s => Array.from(document.querySelectorAll(s));
const main = e => !e.closest('.ant-drawer, .ant-modal-wrap, .el-dialog__wrapper, .el-drawer__wrapper, .ant-picker-dropdown');
const rows = q('.el-table__body-wrapper tbody tr, .ant-table-tbody tr, table tbody tr').filter(e => vis(e) && main(e));
const head = rows.slice(0, 3).map(r => r.textContent).join('|').slice(0, 600);
const empty = q('.el-table__empty-block, .el-table__empty-text, .ant-empty, .ant-table-placeholder').some(e => vis(e) && main(e));
const total = q('.el-pagination__total, .ant-pagination-total-text').filter(e => vis(e) && main(e)).map(e => e.textContent).join('|');
const y = window.__yz;
return {spin: q(arguments[0]).some(vis), sig: rows.length + '#' + head + '#' + total, empty: empty,
        rows: rows.length, net: y ? {p: y.pending, s: y.started} : null,
        msg: q('.el-message--error, .el-message--warning, .ant-message-error, .ant-message-warning, '
               + '.ant-notification-notice-error, .el-notification')
             .filter(vis).map(e => (e.innerText || '').trim()).filter(Boolean).join(' | ').slice(0, 200)};
"""

DETAIL_TEXTS = ["详情", "查看详情", "明细", "查看"]
JS_DETAIL_BTNS = JS_VIS + """
const texts = arguments[0], out = [], keys = new Map();
for (const e of document.querySelectorAll('tbody a, tbody button, tbody span, tbody .ant-btn, tbody .el-button, '
                                          + 'tbody [role=button]')) {
  if (!texts.includes(norm(e.textContent)) || !vis(e)) continue;
  if (e.closest('.el-dialog__wrapper, .ant-modal-wrap, .el-drawer__wrapper, .ant-drawer, [role=dialog]')) continue;
  if (out.some(o => o.contains(e))) continue;
  const tr = e.closest('tr');
  const key = tr ? (tr.getAttribute('data-row-key') || ('#' + Array.from(tr.parentNode.children).indexOf(tr)))
                 : ('@' + out.length);
  // 固定列会把同一行再渲染一份：同一行只保留一个（后出现的那份浮在上层，点得到）
  if (keys.has(key)) { out[keys.get(key)] = e; continue; }
  keys.set(key, out.length);
  out.push(e);
}
return out;
"""

JS_NEXT_PAGE = JS_VIS + """
for (const b of document.querySelectorAll('.el-pagination .btn-next, .ant-pagination-next')) {
  if (!vis(b) || b.closest('.el-dialog__wrapper, .ant-modal-wrap, .ant-drawer, .el-drawer__wrapper')) continue;
  if (b.disabled || b.classList.contains('ant-pagination-disabled') || b.getAttribute('aria-disabled') === 'true') return false;
  return b;
}
return false;
"""

JS_TOP_DIALOG = JS_VIS + """
const ds = Array.from(document.querySelectorAll(arguments[0])).filter(vis)
  .filter(d => !d.classList.contains('el-message-box__wrapper'));
if (!ds.length) return null;
let best = ds[0], bz = -1;
for (const d of ds) { const z = parseInt(getComputedStyle(d).zIndex) || 0; if (z > bz) { bz = z; best = d; } }
return best;
"""
JS_DIALOG_COUNT = JS_VIS + """
return Array.from(document.querySelectorAll(arguments[0])).filter(vis)
  .filter(d => !d.classList.contains('el-message-box__wrapper')).length;
"""


def find_button(driver, scope, exact, part=(), skip_table=True):
    r = _js(driver, JS_FIND_BTN, scope, list(exact), list(part), skip_table)
    return r if r else None


def click_button(driver, texts, timeout):
    r = wait_until(lambda: (lambda x: x if x and x[1] else None)(find_button(driver, None, texts)),
                   timeout, poll=0.3)
    if not r:
        return False
    safe_click(driver, r[0])
    return True


def table_state(driver):
    return _js(driver, JS_T_STATE, SPIN_CSS, default={}) or {}


def goto_next_page(driver):
    btn = _js(driver, JS_NEXT_PAGE, default=False)
    if not btn:
        return False
    before = table_state(driver).get("sig")
    net_reset(driver)
    safe_click(driver, btn)
    wait_until(lambda: table_state(driver).get("sig") != before, 15, poll=0.3)
    wait_until(lambda: not (table_state(driver).get("net") or {}).get("p") and not spinning(driver), 30, poll=0.3)
    return True


# =====================================================================
# 下载
# =====================================================================
DL_EXACT = ["导出", "导出明细", "导出详情", "导出Excel", "导出excel", "下载", "下载详情", "下载明细", "下载Excel"]
DL_PART = ["导出", "下载"]
EXPORT_MSG_CSS = (".ant-message-notice, .ant-notification-notice, .ant-modal-confirm-content, "
                  ".el-message, .el-notification, .el-message-box__message")


def _dirs():
    return watch_dirs(CONFIG["DOWNLOAD_DIR"])


def _files():
    return list_files(_dirs())


def wait_download(driver, before, timeout, idle_fail=None):
    """等待新文件出现。idle_fail：多少秒内既无文件也无网络/下载动静就返回 None（交给上层重点「导出」）。"""
    t0 = time.time()
    end = t0 + timeout
    last_log, last_act, net_sig, announced = t0, t0, None, False
    seen_msgs = set()
    while time.time() < end:
        check_stop()
        now = time.time()
        if now - t0 < 15:
            confirm_msgbox(driver)
        m = _js(driver, JS_VISIBLE_TEXT, EXPORT_MSG_CSS, default="") or ""
        if m and m not in seen_msgs:
            seen_msgs.add(m)
            log(f"页面提示：{m}")
            if re.search(r"下载中心|导出任务|任务列表|稍后.*下载", m):
                raise RuntimeError(f"导出改为后台任务（{m}），需要到页面的下载中心取文件，请把截图发给维护人员适配")
            if re.search(r"失败|错误|异常|无数据|没有数据", m):
                raise RuntimeError(f"导出时页面提示：{m}")
        new = _files() - before
        if new:
            p = max(new, key=lambda x: os.path.getmtime(x))
            if file_stable(p):
                return p
            last_act = now
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
                if st.get("p", 0) > 0:
                    last_act = now
            if idle_fail and now - last_act >= idle_fail:
                return None
        if now - last_log >= 20:
            log(f"等待下载文件 ... 已等待 {fmt_secs(now - t0)}")
            last_log = now
        time.sleep(0.4)
    return None


JS_DIALOG_GONE = JS_VIS + """
const d = arguments[0];
if (!d || !document.body.contains(d)) return true;
if (d.classList.contains('ant-drawer')) return !d.classList.contains('ant-drawer-open');
return !vis(d);
"""


def _close_dialog(driver, dlg):
    def gone():
        try:
            return bool(_js(driver, JS_DIALOG_GONE, dlg, default=True))
        except (StaleElementReferenceException, NoSuchElementException):
            return True
    try:
        for sel in (".el-dialog__headerbtn", ".el-drawer__close-btn", ".ant-modal-close", ".ant-drawer-close",
                    "[aria-label='Close']", "[aria-label='close']"):
            for b in dlg.find_elements(By.CSS_SELECTOR, sel):
                if b.is_displayed():
                    safe_click(driver, b)
                    if wait_until(gone, 3, poll=0.2):
                        return
        r = find_button(driver, dlg, ["关闭", "取消", "返回"], skip_table=False)
        if r:
            safe_click(driver, r[0])
            if wait_until(gone, 3, poll=0.2):
                return
        driver.switch_to.active_element.send_keys(Keys.ESCAPE)
        wait_until(gone, 3, poll=0.2)
    except (StaleElementReferenceException, NoSuchElementException):
        pass
    except WebDriverException as e:
        _raise_if_dead(e)


def _close_new_windows(driver, before):
    for h in [h for h in _handles(driver) if h not in before and h != _WIN.get("page")]:
        try:
            driver.switch_to.window(h)
            driver.close()
        except WebDriverException as e:
            _raise_if_dead(e)
    ensure_page(driver)


def _find_dl_any_context(driver):
    for p in _frame_paths(driver):
        if _enter_path(driver, p):
            r = find_button(driver, None, DL_EXACT, DL_PART)
            if r:
                return r
    return None


def download_detail(driver, btn):
    """点一行的「详情」→ 弹窗里点「导出」→ 返回下载到的文件路径。"""
    before_files = _files()
    handles_before = set(_handles(driver))
    n_dlg = _js(driver, JS_DIALOG_COUNT, DIALOG_CSS, default=0) or 0
    url0 = _js(driver, "return location.href;", default="")
    net_reset(driver)
    safe_click(driver, btn)

    def opened():
        hs = [h for h in _handles(driver) if h not in handles_before]
        if hs:
            return ("win", hs[-1])
        if (_js(driver, JS_DIALOG_COUNT, DIALOG_CSS, default=0) or 0) > n_dlg:
            return ("dlg", None)
        if _files() - before_files:
            return ("file", None)
        if _js(driver, "return location.href;", default=url0) != url0:
            return ("route", None)
        return None

    r = wait_until(opened, CONFIG["DETAIL_TIMEOUT"], poll=0.3)
    if r is None:
        if _js(driver, JS_TOP_DIALOG, DIALOG_CSS):
            r = ("dlg", None)
        else:
            save_debug(driver, "详情未打开")
            raise RuntimeError("点击「详情」后没有弹出窗口")
    kind = r[0]
    if kind == "file":
        p = wait_download(driver, before_files, 30)
        if p:
            return p

    scope = None
    if kind == "win":
        driver.switch_to.window(r[1])
        log("详情在新窗口中打开")
    elif kind == "dlg":
        scope = wait_until(lambda: _js(driver, JS_TOP_DIALOG, DIALOG_CSS), 5, poll=0.2)

    path = None
    try:
        wait_until(lambda: not spinning(driver, scope)
                   and not (net_state(driver) or {}).get("p"), CONFIG["DETAIL_TIMEOUT"], poll=0.3)

        def dl_btn():
            x = _find_dl_any_context(driver) if kind == "win" else find_button(driver, scope, DL_EXACT, DL_PART)
            return x if x and x[1] else None

        b = wait_until(dl_btn, CONFIG["DETAIL_TIMEOUT"], poll=0.4)
        if b is None:
            save_debug(driver, "未找到下载按钮")
            raise RuntimeError("详情窗口中未找到可用的「导出」按钮")
        for attempt in range(2):
            net_hook(driver)
            net_reset(driver)
            safe_click(driver, b[0])
            log("  已点击「导出」" + ("（重试）" if attempt else ""))
            path = wait_download(driver, before_files,
                                 CONFIG["DOWNLOAD_TIMEOUT"] if attempt else CONFIG["DOWNLOAD_IDLE_RETRY"] + 5,
                                 idle_fail=None if attempt else CONFIG["DOWNLOAD_IDLE_RETRY"])
            if path:
                break
            b = dl_btn() or b
        if path is None:
            save_debug(driver, "下载超时")
            raise RuntimeError("点击「导出」后一直没有收到文件")
        return path
    finally:
        try:
            if kind == "win":
                _close_new_windows(driver, handles_before)
            else:
                if kind == "dlg" and scope is not None:
                    _close_dialog(driver, scope)
                elif kind == "route":
                    _js(driver, "history.back();")
                    wait_until(lambda: _js(driver, JS_HAS_FORM, default=False), 15)
                if [h for h in _handles(driver) if h not in handles_before]:
                    _close_new_windows(driver, handles_before)
        except (StopRequested, InvalidSessionIdException):
            raise
        except Exception as e:
            _raise_if_dead(e)
            log(f"关闭详情窗口时出错（忽略）：{str(e)[:60]}")


# =====================================================================
# 命名 / 清单 / 汇总
# =====================================================================
JS_ROW_INFO = JS_VIS + """
const tr = arguments[0].closest('tr'); if (!tr) return null;
const key = tr.getAttribute('data-row-key');
let row = tr;
if (key) {   // 操作列常是右侧固定列：回主表取同一行（单元格最多的那份）
  const rows = Array.from(document.querySelectorAll('tr[data-row-key]')).filter(r => r.getAttribute('data-row-key') === key);
  rows.sort((a, b) => b.cells.length - a.cells.length);
  row = rows[0] || tr;
}
const root = row.closest('.ant-table, .el-table') || row.closest('table');
const heads = Array.from(root.querySelectorAll('thead th')).filter(th => !th.classList.contains('ant-table-cell-scrollbar'))
  .map(th => (th.textContent || '').trim());
const cells = Array.from(row.children).filter(td => td.tagName === 'TD');
const out = {};
cells.forEach((td, i) => { const h = heads[i] || ('列' + (i + 1)); if (h === '操作' || !h) return; out[h] = (td.textContent || '').trim(); });
return out;
"""

INDEX_NAME = "转移记录清单.csv"
SUMMARY_NAME = "汇总_转移记录与档案明细.xlsx"


def period_tag(start, end):
    """2026-06-01 ~ 2026-07-31 → 202606-202607；同一个月 → 202606；非整月 → 20260605-20260720"""
    try:
        a = datetime.date.fromisoformat(norm_datetime(start)[:10])
        b = datetime.date.fromisoformat(norm_datetime(end, True)[:10])
    except ValueError:
        return f"{_digits(start)[:8]}-{_digits(end)[:8]}"
    month_end = (b.replace(day=28) + datetime.timedelta(days=4)).replace(day=1) - datetime.timedelta(days=1)
    if a.day == 1 and b == month_end:
        x, y = a.strftime("%Y%m"), b.strftime("%Y%m")
        return x if x == y else f"{x}-{y}"
    return f"{a:%Y%m%d}-{b:%Y%m%d}"


def record_name(info, n):
    """单据编号_转移类型_转移方式_转出场-转入场_猪只种类_数量头_转移日期"""
    g = lambda k: (str(info.get(k) or "").strip().replace("--", "") or "")
    no = g("单据编号") or f"第{n:04d}条"
    route = f"{g('转出场') or '无'}-{g('转入场') or '无'}"
    kinds = re.sub(r"\s*[,，、;；/]\s*", "+", g("猪只种类"))[:40]
    qty = g("总头数") or g("数量") or g("销售数量")
    qty = f"{_digits(qty) or qty}头" if qty else ""
    day = _digits(g("转移时间"))[:8]
    parts = [no, g("转移类型"), g("转移方式"), route, kinds, qty, day]
    return "_".join(sanitize(p) for p in parts if p)[:180]


# ---------------------------------------------------------------- 已下完的整月：跳过
DONE_MARK = ".下载完成.json"     # 一个任务完整下载成功（没有失败、条数齐全）后写入任务文件夹


def task_folder(group):
    """不打开页面，按任务条件算出它的保存文件夹（与 download_group 的命名一致）。场不确定（当前场）时返回 None。"""
    f = group.get("filters", {})
    farm = str(f.get(FARM_FIELD) or "").strip()
    if farm in KEEP_FARM or not all(f.get(k) for k in TIME_FIELDS):
        return None
    start, end = f["起始转移时间"], f["截止转移时间"]
    base = core.period_folder(CONFIG["DOWNLOAD_DIR"], norm_datetime(start)[:10], norm_datetime(end, True)[:10])
    return os.path.join(base, sanitize(f"{TITLE}_{farm}_{period_tag(start, end)}"))


def skip_info(group, today=None):
    """整月（或连续几个整月）的任务、以前完整下载过、而且不在最近 N 个月内 → 返回 (文件夹, 条数)，可以跳过。
    没有完成标记（中途中断、以前的版本下载的）一律不跳过，会照常搜索，已导出的单据仍会自动跳过。"""
    if CONFIG.get("MANUAL_PARAMS"):
        return None                      # 手动设置了其他条件时，同名文件夹里的数据不一定是同一批
    f = group.get("filters", {})
    try:
        a = datetime.date.fromisoformat(norm_datetime(f.get("起始转移时间"))[:10])
        b = datetime.date.fromisoformat(norm_datetime(f.get("截止转移时间"), True)[:10])
    except (ValueError, TypeError):
        return None
    month_end = (b.replace(day=28) + datetime.timedelta(days=4)).replace(day=1) - datetime.timedelta(days=1)
    if a.day != 1 or b != month_end:
        return None                      # 不是整月（本月至今等）：总是重新下载
    today = today or datetime.date.today()
    if (today.year * 12 + today.month) - (b.year * 12 + b.month) <= CONFIG["REFRESH_RECENT_MONTHS"]:
        return None                      # 最近的月份（含当月）：照常检查，可能有补录
    folder = task_folder(group)
    mark = os.path.join(folder, DONE_MARK) if folder else None
    if not mark or not os.path.isfile(mark):
        return None
    try:
        import json
        with open(mark, encoding="utf-8") as fh:
            rows = json.load(fh).get("rows")
    except (OSError, ValueError):
        rows = None
    return folder, rows


def _write_done_mark(folder, total, rows):
    try:
        import json
        with open(os.path.join(folder, DONE_MARK), "w", encoding="utf-8") as fh:
            json.dump({"total": total, "rows": rows,
                       "time": datetime.datetime.now().isoformat(timespec="seconds")}, fh, ensure_ascii=False)
    except OSError:
        pass


def _existing_record(folder, no):
    if not no:
        return None
    try:
        for f in os.listdir(folder):
            if (f.startswith(sanitize(no) + "_") and f not in (INDEX_NAME, SUMMARY_NAME)
                    and not f.lower().endswith(core.PARTIAL_EXT)):
                return os.path.join(folder, f)
    except OSError:
        pass
    return None


class RecordIndex:
    """转移记录清单：每行 = 一张转移单 + 导出文件名 + 状态。逐条写入 CSV（中途中断不丢），结束时生成汇总 xlsx。"""

    def __init__(self, folder):
        self.path = os.path.join(folder, INDEX_NAME)
        self.rows, self.cols = {}, []
        if os.path.exists(self.path):          # 断点续传：读入已有清单
            try:
                with open(self.path, encoding="utf-8-sig", newline="") as f:
                    r = csv.DictReader(f)
                    self.cols = [c for c in (r.fieldnames or []) if c not in ("导出状态", "导出文件", "导出时间")]
                    for k, row in enumerate(r):
                        self.rows[row.get("单据编号") or row.get("序号") or f"#{k}"] = row
            except Exception:
                self.rows, self.cols = {}, []

    def add(self, info, file, status, n=None):
        for k in info:
            if k not in self.cols:
                self.cols.append(k)
        row = dict(info)
        row.update({"导出状态": status, "导出文件": os.path.basename(file) if file else "",
                    "导出时间": time.strftime("%Y-%m-%d %H:%M:%S")})
        # 旧版没有单据编号的行 key 都是 None，会互相覆盖
        key = info.get("单据编号") or info.get("序号") or f"#{n if n is not None else len(self.rows)}"
        self.rows[key] = row
        self._write()

    def _write(self):
        cols = self.cols + ["导出状态", "导出文件", "导出时间"]
        tmp = self.path + ".tmp"
        for attempt in range(5):
            try:
                with open(tmp, "w", encoding="utf-8-sig", newline="") as f:
                    w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
                    w.writeheader()
                    for row in self.rows.values():
                        w.writerow(row)
                os.replace(tmp, self.path)
                return
            except PermissionError:
                if attempt == 0:
                    log(f"  提示：{INDEX_NAME} 被占用（可能正用 Excel 打开），稍后重试写入")
                time.sleep(1)
        log(f"  警告：{INDEX_NAME} 写入失败（文件被占用），本条记录会在下次写入时补上")

    def build_summary(self, folder):
        """汇总 xlsx：①转移记录（每张单一行）②档案明细（每个导出文件的明细行前加上所属转移单的信息）。"""
        try:
            import openpyxl
        except ImportError:
            log("提示：未安装 openpyxl，跳过生成汇总表（转移记录清单.csv 已生成）")
            return None
        rec_cols = self.cols
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "转移记录"
        ws.append(rec_cols + ["导出状态", "导出文件"])
        for row in self.rows.values():
            ws.append([row.get(c, "") for c in rec_cols] + [row.get("导出状态", ""), row.get("导出文件", "")])

        ws2 = wb.create_sheet("档案明细")
        detail_cols, detail_rows, skipped = [], [], 0
        for row in self.rows.values():
            fn = row.get("导出文件")
            if not fn or row.get("导出状态") not in ("完成", "已存在"):
                continue
            head, data = _read_table(os.path.join(folder, fn))
            if head is None:
                skipped += 1
                continue
            for h in head:
                if h not in detail_cols:
                    detail_cols.append(h)
            for d in data:
                detail_rows.append((row, dict(zip(head, d))))
        link_cols = [c for c in ("单据编号", "转移方式", "转移类型", "转移时间", "转出场", "转入场") if c in rec_cols]
        ws2.append(["所属" + c if c in detail_cols else c for c in link_cols] + detail_cols)
        for rec, d in detail_rows:
            ws2.append([rec.get(c, "") for c in link_cols] + [d.get(c, "") for c in detail_cols])
        for sh in (ws, ws2):
            sh.freeze_panes = "A2"
            for col in sh.columns:
                width = max((len(str(c.value or "")) for c in list(col)[:200]), default=4)
                sh.column_dimensions[col[0].column_letter].width = min(max(width * 1.6, 6), 40)
        out = os.path.join(folder, SUMMARY_NAME)
        try:
            wb.save(out)
        except PermissionError:
            out = os.path.join(folder, time.strftime("汇总_%Y%m%d_%H%M%S.xlsx"))
            wb.save(out)
        log(f"已生成汇总表：{os.path.basename(out)}（{len(self.rows)} 张转移单，{len(detail_rows)} 行档案明细"
            + (f"，{skipped} 个文件无法读取" if skipped else "") + "）")
        return out


def _read_table(path):
    """读取导出文件的第一个表：表头 = 第一行至少 3 个非空单元格的行。"""
    rows = []
    try:
        ext = os.path.splitext(path)[1].lower()
        if ext in (".xlsx", ".xlsm"):
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            rows = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
            wb.close()
        elif ext == ".xls":
            import xlrd
            sh = xlrd.open_workbook(path).sheet_by_index(0)
            rows = [sh.row_values(i) for i in range(sh.nrows)]
        elif ext == ".csv":
            with open(path, encoding="utf-8-sig", errors="replace", newline="") as f:
                rows = list(csv.reader(f))
        else:
            return None, None
    except Exception:
        return None, None
    for i, r in enumerate(rows):
        if sum(1 for c in r if str(c or "").strip()) >= 3:
            head = [str(c or "").strip() or f"列{j + 1}" for j, c in enumerate(r)]
            data = [r2 for r2 in rows[i + 1:] if any(str(c or "").strip() for c in r2)]
            return head, data
    return None, None


JS_CUR_PAGE = JS_VIS + """
for (const li of document.querySelectorAll('.ant-pagination-item-active, .el-pager .number.active, .el-pager li.active')) {
  if (!vis(li) || li.closest('.ant-drawer, .ant-modal-wrap, .el-dialog__wrapper, .el-drawer__wrapper')) continue;
  const n = parseInt(li.getAttribute('title') || li.textContent, 10);
  if (!isNaN(n)) return n;
}
return null;
"""
JS_PAGE_BTN = JS_VIS + """
const want = String(arguments[0]), way = arguments[1];
const main = e => !e.closest('.ant-drawer, .ant-modal-wrap, .el-dialog__wrapper, .el-drawer__wrapper');
const item = Array.from(document.querySelectorAll('.ant-pagination-item, .el-pager .number'))
  .find(li => vis(li) && main(li) && (li.getAttribute('title') || li.textContent.trim()) === want);
if (item) { (item.querySelector('a') || item).click(); return 'item'; }
const sel = way === 'next' ? '.ant-pagination-next, .el-pagination .btn-next' : '.ant-pagination-prev, .el-pagination .btn-prev';
const b = Array.from(document.querySelectorAll(sel)).find(x => vis(x) && main(x)
  && !x.classList.contains('ant-pagination-disabled') && !x.disabled);
if (b) { (b.querySelector('button') || b).click(); return way; }
return null;
"""


def _ensure_page_no(driver, page):
    """关闭详情后列表若跳回了别的页，翻回第 page 页。"""
    for _ in range(60):
        cur = _js(driver, JS_CUR_PAGE)
        if cur is None or cur == page:
            return True
        r = _js(driver, JS_PAGE_BTN, page, "next" if cur < page else "prev")
        if not r:
            return False
        wait_until(lambda: _js(driver, JS_CUR_PAGE) != cur, 5, poll=0.2)
        _wait_rows(driver, 1, 15)
    return False


def _wait_rows(driver, need, timeout):
    def ok():
        if spinning(driver):
            return False
        return len(_js(driver, JS_DETAIL_BTNS, DETAIL_TEXTS) or []) >= need
    return bool(wait_until(ok, timeout, poll=0.3))


# =====================================================================
# 单个任务
# =====================================================================
def download_one_group(driver, group):
    """一个任务 = 一个场 + 一个转移时间段。返回 {"path": 文件夹, "total", "rows"}；搜索无数据返回 None。"""
    name = group["name"]
    filters = normalize_filters(group.get("filters", {}))
    t_start = time.time()
    log(f"---- 开始任务: {name} ----")
    ensure_page(driver)
    close_popups(driver)
    manual = bool(CONFIG.get("MANUAL_PARAMS"))

    if not manual:
        log(f"设置: {FARM_FIELD} = {filters.get(FARM_FIELD) or '当前场'}")
        set_farm(driver, filters.get(FARM_FIELD, ""))
        if CONFIG.get("RESET_BETWEEN_TASKS") and click_button(driver, ["重置"], 2):
            wait_loading_done(driver, 15)
            time.sleep(0.3)
        expand_filters(driver)
        for field in [f for f in SELECT_FIELDS + TEXT_FIELDS if str(filters.get(f) or "").strip() not in ALL_VALUES]:
            log(f"设置: {field} = {filters[field]}")
            apply_filter(driver, field, filters[field])
    if all(f in filters for f in TIME_FIELDS):
        log(f"设置: 转移时间 = {filters['起始转移时间']} ~ {filters['截止转移时间']}")
        set_date_range(driver, filters["起始转移时间"], filters["截止转移时间"])
    else:
        for field in [f for f in TIME_FIELDS if f in filters]:
            log(f"设置: {field} = {filters[field]}")
            apply_filter(driver, field, filters[field])
    close_popups(driver)

    farm = read_current_params(driver).get("__farm") or str(filters.get(FARM_FIELD) or "").strip()
    if farm in KEEP_FARM:
        farm = "当前场"
    period = period_tag(filters.get("起始转移时间", ""), filters.get("截止转移时间", ""))
    base = core.period_folder(CONFIG["DOWNLOAD_DIR"], norm_datetime(filters["起始转移时间"])[:10],
                              norm_datetime(filters["截止转移时间"], True)[:10]) \
        if all(filters.get(f) for f in TIME_FIELDS) else CONFIG["DOWNLOAD_DIR"]
    folder = os.path.join(base, sanitize(f"{TITLE}_{farm}_{period}"))

    net_hook(driver)
    net_reset(driver)
    before = table_state(driver)
    log("点击「搜索」...")
    if not click_button(driver, ["搜索", "查询"], 15):
        raise RuntimeError("未找到「搜索」按钮")
    wait_query_done(driver, before, time.time(), CONFIG["SEARCH_TIMEOUT"], CONFIG["SEARCH_IDLE_GRACE"],
                    state_fn=table_state, err_re=r"失败|错误|异常|请选择|不能|超出|超过", what="搜索")

    total = query_total(driver)
    btns = wait_until(lambda: _js(driver, JS_DETAIL_BTNS, DETAIL_TEXTS), 5, poll=0.4) or []
    if total is not None:
        log(f"搜索结果共 {total} 条转移记录")
    if not btns:
        if total == 0 or table_state(driver).get("empty"):
            log(f"[无数据] {name}：该场在这个时间段没有转移记录")
            return None
        save_debug(driver, f"无详情按钮_{name}")
        raise RuntimeError("搜索后表格中没有找到「详情」按钮")

    os.makedirs(folder, exist_ok=True)
    try:
        os.remove(os.path.join(folder, DONE_MARK))   # 本次没跑完之前不算「已下完」，成功后再写
    except OSError:
        pass
    index = RecordIndex(folder)
    if index.rows:
        log(f"文件夹里已有 {len(index.rows)} 条记录，已导出的会自动跳过（断点续传）")
    n = done = skipped = failed = streak = 0
    page = 1
    seen = set()
    core.progress(0, total)
    if _js(driver, JS_CUR_PAGE) not in (None, 1):
        log(f"  列表停在第 {_js(driver, JS_CUR_PAGE)} 页，先回到第 1 页")
        _ensure_page_no(driver, 1)
    while True:
        _wait_rows(driver, 1, 15)
        count = len(_js(driver, JS_DETAIL_BTNS, DETAIL_TEXTS) or [])
        for k in range(count):
            check_stop()
            _wait_rows(driver, k + 1, 20)
            if page > 1 and _js(driver, JS_CUR_PAGE) not in (None, page):
                log(f"  列表跳回了第 {_js(driver, JS_CUR_PAGE)} 页，翻回第 {page} 页")
                _ensure_page_no(driver, page)
                _wait_rows(driver, k + 1, 20)
            cur = _js(driver, JS_DETAIL_BTNS, DETAIL_TEXTS) or []
            if k >= len(cur):
                log(f"  警告：列表只剩 {len(cur)} 行（原 {count} 行），本页后面的记录稍后核对补齐")
                break
            info = _js(driver, JS_ROW_INFO, cur[k], default={}) or {}
            no = info.get("单据编号", "")
            if no and no in seen:
                continue
            if no:
                seen.add(no)
            n += 1
            core.progress(n, total)
            label = f"第 {n}/{total or '?'} 条 {no} {info.get('转移类型', '')} {info.get('转移时间', '')}".strip()
            old = _existing_record(folder, no)
            if old:
                skipped += 1
                if (index.rows.get(no) or {}).get("导出状态") not in ("完成", "已存在"):
                    index.add(info, old, "已存在", n)
                continue
            log(f"== {label}（第 {page} 页）==")
            try:
                path = download_detail(driver, cur[k])
                final = move_unique(path, folder, record_name(info, n))
                index.add(info, final, "完成", n)
                done += 1
                streak = 0
                log(f"  已保存 {os.path.basename(final)}")
            except (StopRequested, InvalidSessionIdException):
                raise
            except Exception as e:
                _raise_if_dead(e)
                failed += 1
                streak += 1
                index.add(info, "", f"失败：{str(e)[:60]}", n)
                log(f"  导出失败：{str(e)[:120]}")
                try:
                    ensure_page(driver)
                    dlg = _js(driver, JS_TOP_DIALOG, DIALOG_CSS)
                    if dlg is not None:
                        _close_dialog(driver, dlg)
                except (StopRequested, InvalidSessionIdException):
                    raise
                except Exception:
                    pass
                if streak >= 3:
                    raise RuntimeError(f"连续 {streak} 条导出失败，可能页面状态异常，重新打开页面后续传")
            ensure_page(driver)
            if n >= CONFIG["MAX_ROWS"]:
                break
        if n >= CONFIG["MAX_ROWS"]:
            log(f"已达单任务上限 {CONFIG['MAX_ROWS']} 条，停止")
            break
        _wait_rows(driver, 1, 20)
        if page > 1:
            _ensure_page_no(driver, page)
        if page > 5000 or not goto_next_page(driver):
            break
        page += 1

    index.build_summary(folder)
    if total and n < min(total, CONFIG["MAX_ROWS"]):
        raise RuntimeError(f"只处理了 {n}/{total} 条转移记录（列表刷新或翻页中断），"
                           "将重新打开页面续传，已导出的会自动跳过")
    log(f"[完成] {name}：新导出 {done} 条" + (f"，跳过已有 {skipped} 条" if skipped else "")
        + (f"，失败 {failed} 条" if failed else "") + f" -> {folder}，用时 {fmt_secs(time.time() - t_start)}")
    if failed:
        raise RuntimeError(f"有 {failed} 条导出失败（已导出的已保存，重试时会自动跳过）")
    if not (n >= CONFIG["MAX_ROWS"] and total and total > n):
        _write_done_mark(folder, total, n)         # 完整下完：以后这个整月可以直接跳过
    return {"path": folder, "total": total, "rows": n}


# =====================================================================
# 会话 / 主流程
# =====================================================================
MANUAL_MSG = ("已进入「猪只转移」页面。请在浏览器里：\n"
              "  1. 右上角选好场；\n"
              "  2. 设置转移类型、转移方式、转入场、转出场、猪只种类等参数（转移时间不用填）；\n"
              "然后回到程序继续。程序会逐个填转移时间、搜索并导出。")


def _goto(driver, manual=None):
    manual = CONFIG.get("MANUAL_PARAMS") if manual is None else manual
    _WIN.update(page=None, frame=[])
    _FIELD_CACHE.clear()
    goto_transfer_page(driver)
    if manual:
        wait_manual(driver, MANUAL_MSG)


def open_session(username, password, manual=None):
    manual = CONFIG.get("MANUAL_PARAMS") if manual is None else manual
    if manual and core.HEADLESS:
        log("手动设置参数需要看到浏览器，本次改为显示浏览器窗口")
        core.HEADLESS = False
    return core.start_session(CONFIG["DOWNLOAD_DIR"], username, password, lambda d: _goto(d, manual))


def reload_page(driver):
    core.back_to_home(driver, _WIN.get("page"))
    if CONFIG.get("MANUAL_PARAMS"):
        log("页面已重新打开，之前手动设置的参数丢失了")
    _goto(driver)


ADAPTER = core.Adapter(goto=_goto, run=download_one_group, reload=reload_page, name=lambda g: g["name"],
                       download_dir=lambda: CONFIG["DOWNLOAD_DIR"], title=TITLE,
                       open_session=open_session, begin=_FIELD_CACHE.clear)


def run_groups(groups, username=None, password=None, status_fn=None):
    return core.run_tasks(groups, ADAPTER, username, password, status_fn, what="任务")


def fetch_options(username=None, password=None, fields=None):
    """登录并读取各下拉字段的可选项，返回 {字段: [选项, ...]}。"""
    if not username or not password:
        username, password = core.load_credentials()
    out = {}
    try:
        driver = open_session(username, password, manual=False)
        ensure_page(driver)
        expand_filters(driver)
        for f in fields or PARAM_FIELDS:
            check_stop()
            try:
                opts = read_options(driver, f)
                out[f] = opts
                log(f"「{f}」共 {len(opts)} 个选项：{'、'.join(opts[:15])}{' …' if len(opts) > 15 else ''}")
            except InvalidSessionIdException:
                raise
            except Exception as e:
                _raise_if_dead(e)
                log(f"读取「{f}」选项失败：{str(e)[:100]}")
    finally:
        core.force_quit()
    return out


def main():
    if "--headless" in sys.argv:
        core.HEADLESS = True
    username, password = core.load_credentials()
    if "--options" in sys.argv:
        fetch_options(username, password)
        return
    groups = CONFIG["FILTER_GROUPS"]
    for g in groups:
        f = normalize_filters(g.get("filters", {}))
        g["name"] = g.get("name") or f"{f.get(FARM_FIELD, '当前场')}_" \
                                     f"{period_tag(f.get('起始转移时间', ''), f.get('截止转移时间', ''))}"
    run_groups(groups, username, password)


if __name__ == "__main__":
    main()
