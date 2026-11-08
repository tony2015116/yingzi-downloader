# -*- coding: utf-8 -*-
"""
系谱查询（基因组选种平台 · 基础信息 › 系谱信息）—— 按国标编码批量导出系谱
================================================================
首页「基因组选种」→ 左侧「基础信息」→「系谱信息」→ 标签「系谱信息」。
每次查询用三个条件：右上角的「场」、查询区的「在场 / 离场」、「国标编码」（点「批量」粘贴，一次最多 400 个）。

一批编码里的猪可能分布在不同的场，有的已经离场，一次查不全，所以循环查询：
    先「在场」，按场列表逐个场查；查到的编码从待查名单里去掉，剩下的再查下一个场；
    所有场都查过后，还没找到的再按「离场」逐个场查一遍。
    全部找到就提前结束。
每次查询导出的文件都按参数命名保留，最后合并成一个完整的文件，并给出每个编码在哪个场 / 状态找到、哪些没找到。

文件：系谱查询\\2026年\\10月\\系谱_国标N个_{年月日_时分秒}\\
        01_龙马一场_在场_第1批_400个编码_312条.xlsx      ← 每次查询导出的原始文件
        …
        系谱_国标N个_找到M个_{R}条_合并.xlsx              ← 合并后的完整文件（加了「查询场」「查询状态」两列）
        查询明细.csv                                        ← 每个编码：在哪个场 / 状态找到，或未找到
        未找到的国标编码.txt                                ← 有未找到的才生成

    python task_pedigree.py --codes 国标编码.txt --farms "龙马一场;龙马二场" [--status 在场|离场|全部] [--headless]
"""

import os
import re
import sys
import csv
import time
import shutil
import datetime
from collections import ChainMap

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import InvalidSessionIdException, WebDriverException, StaleElementReferenceException

import yz_core as core
import task_genetic as gn
import task_ebv as ev
from yz_core import (StopRequested, log, wait_until, fmt_secs, sanitize, safe_click, _js, _raise_if_dead,
                     _handles, _current_handle, net_hook, net_reset, net_dump, wait_loading_done, click_text,
                     find_menu, switch_to_new_window, save_debug, open_url, page_snapshot, watch_dirs, list_files,
                     move_replace, period_folder, rel_to, xlsx_data_rows, append_csv)

TITLE = "系谱查询"

CONFIG = ChainMap({
    "PED_URL": "https://gsm.yingzi.com/#/base-info/genealogy",      # 菜单点不到时直接打开
    "DOWNLOAD_DIR": r"D:\test\downloads\系谱查询",
    "TAB": "系谱信息",
    "TAB_KEY": "1",
    "CODE_INPUT_ID": "qp-personId-in",
    "STATUS_ID": "qp-mngStatus-eq",
}, gn.CONFIG)          # 平台地址、等待时间、每批 400 个等与遗传性状查询共用

STATUS_ORDER = {"在场优先，再查离场": ["在场", "离场"], "只查在场": ["在场"], "只查离场": ["离场"]}
STATUS_DEFAULT = "在场优先，再查离场"
FILE_PREFIX = "系谱"
RECORD_NAME = f"{TITLE}下载记录.csv"
_MAIN = {"handle": None}


# =====================================================================
# 任务 / 命名
# =====================================================================
def make_task(codes, farms, status=None):
    codes, _bad = gn.parse_codes("\n".join(codes or []))            # 去重（不区分大小写），保持粘贴顺序
    farms = [f.strip() for f in (farms or []) if f and f.strip()]
    farms = list(dict.fromkeys(farms))
    status = status if status in STATUS_ORDER else STATUS_DEFAULT
    return {"codes": codes, "farms": farms, "status": status}


def task_name(task):
    c, f = task["codes"], task["farms"]
    fs = "、".join(f[:3]) + (f" 等 {len(f)} 个场" if len(f) > 3 else "")
    return f"国标 {len(c)} 个 · {fs or '（未选场）'} · {task.get('status') or STATUS_DEFAULT}"


def make_stem(task):
    return sanitize(f"{FILE_PREFIX}_国标{len(task['codes'])}个")


# =====================================================================
# 导航：首页「基因组选种」→ 基础信息 → 系谱信息
# =====================================================================
JS_ON_PAGE = core.JS_VIS + """
const title = Array.from(document.querySelectorAll('.ant-page-header-heading-title')).filter(vis).map(e => norm(e.textContent));
const tabs = Array.from(document.querySelectorAll('.ant-tabs-tab')).filter(vis).map(t => norm(t.textContent));
return title.includes('系谱信息') && tabs.includes(arguments[0]);
"""


def on_page(driver):
    return CONFIG["GSM_HOST"] in gn._url(driver) and bool(_js(driver, JS_ON_PAGE, CONFIG["TAB"], default=False))


def goto_page(driver, max_attempts=3):
    for attempt in range(1, max_attempts + 1):
        try:
            try:
                driver.switch_to.default_content()
            except WebDriverException as e:
                _raise_if_dead(e)
            if not on_page(driver):
                if CONFIG["GSM_HOST"] not in gn._url(driver):
                    log(f"进入「{CONFIG['PORTAL_TILE']}」...")
                    before = set(_handles(driver))
                    try:
                        click_text(driver, CONFIG["PORTAL_TILE"], timeout=15)
                        switch_to_new_window(driver, before, wait=4)
                        wait_loading_done(driver, timeout=60)
                        wait_until(lambda: CONFIG["GSM_HOST"] in gn._url(driver), 20)
                    except (StopRequested, InvalidSessionIdException):
                        raise
                    except Exception as e:
                        _raise_if_dead(e)
                        log(f"未能通过首页「{CONFIG['PORTAL_TILE']}」进入（{str(e)[:60]}），直接打开系谱信息页面")
                        open_url(driver, CONFIG["PED_URL"])
                if CONFIG["GSM_HOST"] in gn._url(driver) and not on_page(driver):
                    item = wait_until(lambda: find_menu(driver, "系谱信息"), 3)
                    if item is None:
                        grp = wait_until(lambda: find_menu(driver, "基础信息"), 8)
                        if grp is not None:
                            log("展开左侧「基础信息」...")
                            safe_click(driver, grp)
                            item = wait_until(lambda: find_menu(driver, "系谱信息"), 5)
                    if item is not None:
                        log("打开「系谱信息」...")
                        safe_click(driver, item)
                    if not wait_until(lambda: on_page(driver), 15):
                        log("菜单点击后未看到系谱信息页面，直接打开页面地址")
                        open_url(driver, CONFIG["PED_URL"])
            if gn._login_lost(driver):
                raise RuntimeError("打开基因组选种平台时跳回了登录页（登录已失效或没有该平台权限）")
            if wait_until(lambda: on_page(driver), CONFIG["TIMEOUT"]):
                wait_loading_done(driver, timeout=30)
                net_hook(driver)
                _MAIN["handle"] = _current_handle(driver)
                log("已进入「系谱信息」页面")
                return True
            log("未能进入系谱信息页面，准备重试 ...")
        except (StopRequested, InvalidSessionIdException):
            raise
        except Exception as e:
            _raise_if_dead(e)
            log(f"导航第 {attempt} 次尝试异常：{type(e).__name__}: {str(e)[:100]}")
        time.sleep(2)
    save_debug(driver, "进入系谱信息失败")
    raise RuntimeError("多次尝试后仍未能进入「系谱信息」页面，请确认账号有「基因组选种」平台的权限。")


def ensure_page(driver):
    mh = _MAIN.get("handle")
    if mh and _current_handle(driver) != mh and mh in _handles(driver):
        driver.switch_to.window(mh)
    if not on_page(driver):
        goto_page(driver)


def select_tab(driver):
    tab, key = CONFIG["TAB"], CONFIG["TAB_KEY"]
    info = wait_until(lambda: _js(driver, gn.JS_TRAIT_TAB, tab, key), 15)
    if not info:
        raise RuntimeError(f"页面上未找到「{tab}」标签")
    if info.get("active"):
        return
    log(f"切换到「{tab}」")
    safe_click(driver, info["el"])
    if not wait_until(lambda: (_js(driver, gn.JS_TRAIT_TAB, tab, key) or {}).get("active"), 6, poll=0.2):
        raise RuntimeError(f"点击「{tab}」后标签未切换")
    wait_loading_done(driver, timeout=60)


# =====================================================================
# 条件：右上角的场、在场 / 离场
# =====================================================================
# 右上角（页面标题右侧）的场选择框
JS_HEADER_SELECT = core.JS_VIS + """
for (const s of document.querySelectorAll('.ant-page-header-heading-extra .ant-select, .ant-pro-page-container .ant-page-header .ant-select')) {
  if (!vis(s)) continue;
  const inp = s.querySelector('input');
  const item = s.querySelector('.ant-select-selection-item');
  return {sel: s, inp: inp, id: inp ? inp.id : '', cur: item ? norm(item.getAttribute('title') || item.textContent) : ''};
}
return null;
"""

JS_SELECT_VALUE = core.JS_VIS + """
const inp = document.getElementById(arguments[0]);
if (!inp) return null;
const s = inp.closest('.ant-select');
if (!s) return null;
const item = s.querySelector('.ant-select-selection-item');
return {sel: s, cur: item ? norm(item.getAttribute('title') || item.textContent) : ''};
"""


def _open_dropdown(driver, sel_el, list_id):
    for k in range(3):
        safe_click(driver, sel_el.find_element(By.CSS_SELECTOR, ".ant-select-selector"))
        if wait_until(lambda: (_js(driver, gn.JS_DD_SCROLL, list_id, None) or {}).get("open"), 3, poll=0.2):
            return True
        gn._escape(driver)
    return False


def _pick_option(driver, list_id, sel_el, inp, name, what):
    """选下拉里文字为 name 的选项：
    1) 点开下拉，在输入框里输入名称（和手动操作一样），出现完全一致的选项就点它；
    2) 输入后搜不到（有的下拉按内部编号搜索），清掉文字，在下拉里逐屏滚动查找；
    3) 还找不到就报错，并列出下拉里所有的名称。"""
    if not _open_dropdown(driver, sel_el, list_id):
        save_debug(driver, f"{what}下拉未打开")
        raise RuntimeError(f"「{what}」下拉列表点不开")
    opt = None
    if inp is not None:
        try:
            inp.send_keys(Keys.CONTROL, "a")
            inp.send_keys(name)
        except WebDriverException as e:
            _raise_if_dead(e)
            if isinstance(e, StaleElementReferenceException):
                raise
        opts = wait_until(lambda: _js(driver, gn.JS_SELECT_OPTIONS, list_id, name, True, default=[]), 3, poll=0.3)
        opt = opts[0] if opts else None
        if opt is None:
            try:                                  # 清掉输入的文字，恢复完整列表
                inp.send_keys(Keys.CONTROL, "a")
                inp.send_keys(Keys.BACKSPACE)
                time.sleep(0.6)
            except WebDriverException as e:
                _raise_if_dead(e)
    if opt is None:
        if not (_js(driver, gn.JS_DD_SCROLL, list_id, None) or {}).get("open"):
            _open_dropdown(driver, sel_el, list_id)
        opt = gn._find_in_dropdown(driver, list_id, name)
    if opt is None:
        seen = gn._all_dropdown_texts(driver, list_id)
        save_debug(driver, f"未找到{what}_{name}")
        gn._escape(driver)
        raise RuntimeError(f"「{what}」下拉列表里没有找到「{name}」。列表中的名称：{'、'.join(seen[:40]) or '（无）'}"
                           + (f" 等 {len(seen)} 个" if len(seen) > 40 else ""))
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'nearest'});", opt)
        opt.click()
    except WebDriverException as e:
        _raise_if_dead(e)
        if isinstance(e, StaleElementReferenceException):
            raise
        driver.execute_script("arguments[0].click();", opt)


def _retry_stale(driver, fn, what, tries=3):
    """页面刚加载 / 换场后会重新渲染，拿到的元素可能已失效：重新获取后再试。"""
    for k in range(tries):
        try:
            return fn()
        except StaleElementReferenceException:
            if k == tries - 1:
                raise
            log(f"{what}：页面刚刷新，重新获取后再试 ...")
            gn._escape(driver)
            time.sleep(1.0)
            wait_loading_done(driver, timeout=30)


def current_farm(driver):
    return (_js(driver, JS_HEADER_SELECT) or {}).get("cur") or ""


def set_farm(driver, farm):
    """右上角选场（就是页面右上角那个场选择框）。换场后页面会刷新数据，等加载完。"""
    def once():
        info = wait_until(lambda: _js(driver, JS_HEADER_SELECT), 10)
        if not info:
            save_debug(driver, "系谱_未找到场选择框")
            raise RuntimeError("页面右上角没有找到「场」选择框")
        if info["cur"] == gn._n(farm):
            return False
        log(f"右上角切换场：{info['cur'] or '（空）'} → {farm}")
        _pick_option(driver, info["id"], info["sel"], info["inp"], farm, "场")
        if not wait_until(lambda: current_farm(driver) == gn._n(farm), 8, poll=0.3):
            save_debug(driver, f"系谱_切换场失败_{farm}")
            raise RuntimeError(f"点选「{farm}」后右上角显示为「{current_farm(driver)}」，切换场未生效")
        gn._escape(driver)
        wait_loading_done(driver, timeout=60)
        time.sleep(0.5)
        return True
    wait_loading_done(driver, timeout=30)
    return _retry_stale(driver, once, "切换场")


JS_STATUS_VISIBLE = core.JS_VIS + """
const inp = document.getElementById(arguments[0]);
const s = inp && inp.closest('.ant-select');
return !!(s && vis(s));
"""
# 查询区右侧的展开 / 收起按钮（︾）：折叠时图标带 collapsed
JS_MORE_BTN = core.JS_VIS + """
for (const b of document.querySelectorAll('.more-btn button, .more-btn .ant-btn')) {
  if (!vis(b) || b.closest('.ant-tabs-tabpane:not(.ant-tabs-tabpane-active)')) continue;
  return b;
}
return null;
"""


def expand_form(driver):
    """「在场 / 离场」在查询区的折叠部分里：看不见时先点右侧的 ︾ 展开。"""
    sid = CONFIG["STATUS_ID"]
    if _js(driver, JS_STATUS_VISIBLE, sid, default=False):
        return
    btn = wait_until(lambda: _js(driver, JS_MORE_BTN), 5)
    if btn is None:
        save_debug(driver, "系谱_未找到展开按钮")
        raise RuntimeError("「在场 / 离场」被折叠了，但没有找到查询区右侧的展开按钮（︾）")
    log("展开全部查询条件（︾）")
    safe_click(driver, btn)
    if not wait_until(lambda: _js(driver, JS_STATUS_VISIBLE, sid, default=False), 5, poll=0.2):
        save_debug(driver, "系谱_展开后看不到在场状态")
        raise RuntimeError("点了展开（︾）后仍看不到「在场 / 离场」选择框")


def set_status(driver, status):
    sid = CONFIG["STATUS_ID"]

    def once():
        info = wait_until(lambda: _js(driver, JS_SELECT_VALUE, sid), 10)
        if not info:
            save_debug(driver, "系谱_未找到在场状态")
            raise RuntimeError("查询条件里没有找到「在场 / 离场」选择框")
        if info["cur"] == status:
            return
        expand_form(driver)
        info = _js(driver, JS_SELECT_VALUE, sid) or info
        inp = driver.find_element(By.ID, sid)
        _pick_option(driver, sid, info["sel"], inp, status, "在场状态")
        if not wait_until(lambda: (_js(driver, JS_SELECT_VALUE, sid) or {}).get("cur") == status, 5, poll=0.2):
            raise RuntimeError(f"「在场状态」未能选为「{status}」")
        gn._escape(driver)
        log(f"在场状态：{status}")
    _retry_stale(driver, once, "选择在场状态")


# =====================================================================
# 一次查询：选场 → 在场状态 → 批量粘贴编码 → 查询 → 导出
# =====================================================================
def _dirs():
    return watch_dirs(CONFIG["DOWNLOAD_DIR"])


def export_once(driver, farm, status, codes):
    """返回 (临时文件, 条数, 页面条数)；查不到返回 None。"""
    ensure_page(driver)
    gn._escape(driver)
    select_tab(driver)
    if not wait_until(lambda: _js(driver, ev.JS_CODE_INPUT, CONFIG["CODE_INPUT_ID"], default=False),
                      CONFIG["TIMEOUT"], poll=0.4):
        save_debug(driver, "系谱_未找到国标编码输入框")
        raise RuntimeError("系谱信息页面上一直没有出现「国标编码」输入框（debug 文件夹里有当时的截图）")
    set_farm(driver, farm)
    select_tab(driver)                    # 换场后页面可能刷新，再确认一次标签
    set_status(driver, status)
    gn.set_codes(driver, codes, CONFIG["CODE_INPUT_ID"])
    total = gn.run_query(driver)
    log(f"查询结果：共 {total if total is not None else '（未读取到）'} 条")
    if total == 0:
        return None
    before = list_files(_dirs(), gn.DOWNLOAD_EXTS)
    handles_before = set(_handles(driver))
    t0 = time.time()
    path = None
    for attempt in range(3):
        btn, why = gn.export_ready(driver, total)
        if btn is None:
            save_debug(driver, "导出按钮不可用_系谱")
            raise RuntimeError(f"「导出」不可用：{why}")
        net_hook(driver)
        net_reset(driver)
        safe_click(driver, btn)
        log("已点击「导出」" + (f"（第 {attempt + 1} 次）" if attempt else ""))
        try:
            path = gn.wait_download(driver, before, CONFIG["EXPORT_RETRY_WAIT"], total, dirs=_dirs())
        except gn.NeedSplit as e:
            raise RuntimeError(f"{e}（系谱查询一次最多 400 个编码，不应超限，请检查）")
        if path:
            break
        _close_new_windows(driver, handles_before)
        gn._escape(driver)
        log("未收到下载文件，重试导出 ...")
    if path is None:
        path = gn.wait_download(driver, before, CONFIG["DOWNLOAD_TIMEOUT"], total, dirs=_dirs())
    if path is None:
        for r in net_dump(driver).get("last") or []:
            log(f"  最近请求：{r.get('u')} → {r.get('s')} {r.get('ct') or ''} {r.get('cd') or ''}")
        page_snapshot(driver)
        save_debug(driver, "下载超时_系谱")
        raise RuntimeError("等待下载文件超时，请检查导出是否触发（debug 文件夹里有当时的截图）")
    _close_new_windows(driver, handles_before)
    rows = xlsx_data_rows(path)
    if rows == 0 and total:
        log(f"警告：导出文件读出 0 条，但页面显示 {total} 条，请打开文件核对")
        rows = None
    log(f"导出完成：{rows if rows is not None else total} 条（用时 {fmt_secs(time.time() - t0)}）")
    return path, (rows if rows is not None else total), total


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


# =====================================================================
# 读导出文件：找到了哪些编码
# =====================================================================
_PARENT_RE = re.compile(r"父|母|祖|曾|外")


def _code_column(header):
    """表头里「国标编码」那一列（不是父亲 / 母亲 / 祖辈的国标）。"""
    hs = [gn._n(v) for v in header]
    for i, h in enumerate(hs):
        if h in ("国标编码", "国标号", "国标编号", "个体号"):
            return i
    for i, h in enumerate(hs):
        if "国标" in h and not _PARENT_RE.search(h):
            return i
    return None


def found_codes(path):
    """导出文件里出现的本猪国标编码（大写）。读不出返回 None。"""
    if not str(path).lower().endswith((".xlsx", ".xlsm")):
        return None
    import openpyxl
    he = gn._header_end(path)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        ws.reset_dimensions()
        col, out = None, set()
        for i, r in enumerate(ws.iter_rows(values_only=True), 1):
            if i <= he:
                c = _code_column(r)
                if c is not None:
                    col = c
                continue
            if col is None:
                return None
            if col < len(r) and r[col] not in (None, ""):
                out.add(str(r[col]).strip().upper())
        return out
    finally:
        wb.close()


def merge_with_source(pieces, out):
    """合并各次导出：表头取第一个文件，末尾加「查询场」「查询状态」两列；序号重新编号。返回数据行数。"""
    import openpyxl
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    wb_out = openpyxl.Workbook(write_only=True)
    ws_out = None
    rows, width, renum = 0, None, False
    bold, fill = Font(bold=True), PatternFill("solid", fgColor="F2F4F7")
    for i, p in enumerate(pieces):
        he = gn._header_end(p["path"])
        wb = openpyxl.load_workbook(p["path"], read_only=True, data_only=True)
        try:
            ws = wb.worksheets[0]
            ws.reset_dimensions()
            if ws_out is None:
                ws_out = wb_out.create_sheet(ws.title or "系谱")
            for r_i, r in enumerate(ws.iter_rows(values_only=True), 1):
                r = list(r)
                if r_i <= he:
                    if i == 0:
                        if width is None:
                            width = len(r)
                        extra = ["查询场", "查询状态"] if r_i == he else ["", ""]
                        if r_i == he:
                            renum = bool(r) and gn._n(r[0]) in ("序号", "#", "No.", "NO")
                            for c, v in enumerate(r + extra, 1):
                                ws_out.column_dimensions[get_column_letter(c)].width = \
                                    max(10, min(40, len(str(v or "")) * 2 + 2))
                            try:
                                ws_out.freeze_panes = f"A{he + 1}"
                            except Exception:
                                pass
                        cells = []
                        for v in r + [None] * max(0, (width or 0) - len(r)) + extra:
                            c = WriteOnlyCell(ws_out, value=v)
                            c.font, c.fill = bold, fill
                            c.alignment = Alignment(horizontal="center", vertical="center")
                            cells.append(c)
                        ws_out.append(cells)
                    continue
                if not any(v not in (None, "") for v in r):
                    continue
                rows += 1
                if renum:
                    r[0] = rows
                r = (r + [None] * max(0, (width or 0) - len(r)))[:max(width or len(r), len(r))]
                ws_out.append(r + [p["farm"], p["status"]])
        finally:
            wb.close()
    if ws_out is None:
        ws_out = wb_out.create_sheet("系谱")
    wb_out.save(out)
    return rows


# =====================================================================
# 一个任务：循环查完所有编码
# =====================================================================
def _unique_folder(base, stem):
    stamp = time.strftime("%Y%m%d_%H%M%S")
    cand, k = os.path.join(base, f"{stem}_{stamp}"), 1
    while os.path.exists(cand):
        k += 1
        cand = os.path.join(base, f"{stem}_{stamp}_{k}")
    return cand


def write_record(task, found, rows, queries, path):
    append_csv(os.path.join(CONFIG["DOWNLOAD_DIR"], RECORD_NAME),
               ["下载时间", "国标编码个数", "找到", "未找到", "场", "在场状态", "查询次数", "合并后条数", "保存位置"],
               [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), len(task["codes"]), found,
                len(task["codes"]) - found, "；".join(task["farms"]), task["status"], queries,
                "" if rows is None else rows, rel_to(path, CONFIG["DOWNLOAD_DIR"]) if path else "（无数据）"])


def download_task(driver, task):
    task = make_task(task.get("codes"), task.get("farms"), task.get("status"))
    codes, farms = task["codes"], task["farms"]
    if not codes:
        raise ValueError("系谱查询必须填写国标编码")
    if not farms:
        raise ValueError("系谱查询至少要选一个场")
    statuses = STATUS_ORDER[task["status"]]
    t0 = time.time()
    log(f"---- 开始：{task_name(task)} ----")
    base = period_folder(CONFIG["DOWNLOAD_DIR"], datetime.date.today(), datetime.date.today())
    stem = make_stem(task)
    folder = _unique_folder(base, stem)
    os.makedirs(folder, exist_ok=True)
    remaining = {c.upper(): c for c in codes}        # 待查：大写 → 原样
    where = {}                                       # 大写编码 → (场, 状态)
    pieces, queries, seq = [], 0, 0
    plan = [(st, fm) for st in statuses for fm in farms]
    log(f"国标编码 {len(codes)} 个（已去重）；查询顺序：" + " → ".join(f"{fm}·{st}" for st, fm in plan))
    try:
        for step, (status, farm) in enumerate(plan, 1):
            if not remaining:
                break
            core.progress(step - 1, len(plan), "轮")
            todo = list(remaining.values())
            batches = gn.code_batches(todo, CONFIG["CODE_BATCH"])
            log(f"== {farm} · {status}：待查 {len(todo)} 个" + (f"，分 {len(batches)} 批" if len(batches) > 1 else "")
                + " ==")
            for k, batch in enumerate(batches, 1):
                queries += 1
                res = export_once(driver, farm, status, batch)
                if res is None:
                    log(f"  {farm} · {status} · 第 {k} 批：没有查到")
                    continue
                tmp, n, page = res
                got = found_codes(tmp)
                if got is None:
                    log("  警告：导出文件里没有找到「国标编码」列，无法核对哪些编码已查到；本批视为全部查到")
                    got = {c.upper() for c in batch}
                hit = [c for c in batch if c.upper() in got and c.upper() in remaining]
                seq += 1
                name = sanitize(f"{seq:02d}_{farm}_{status}_第{k}批_{len(batch)}个编码_找到{len(hit)}个_{gn._count_text(n)}")
                dst = move_replace(tmp, folder, name, os.path.splitext(tmp)[1].lower() or ".xlsx")
                pieces.append({"path": dst, "farm": farm, "status": status, "rows": n, "hit": len(hit)})
                for c in hit:
                    where[c.upper()] = (farm, status, os.path.basename(dst))
                    remaining.pop(c.upper(), None)
                log(f"  找到 {len(hit)} 个，还剩 {len(remaining)} 个未找到")
        core.progress(len(plan), len(plan), "轮")

        # 查询明细：每个编码在哪找到
        with open(os.path.join(folder, "查询明细.csv"), "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["序号", "国标编码", "结果", "所在场（查询用）", "在场状态", "来自文件"])
            for i, c in enumerate(codes, 1):
                hit = where.get(c.upper())
                w.writerow([i, c, "找到" if hit else "未找到"] + (list(hit) if hit else ["", "", ""]))
        missing = [c for c in codes if c.upper() not in where]
        if missing:
            with open(os.path.join(folder, "未找到的国标编码.txt"), "w", encoding="utf-8") as f:
                f.write("\n".join(missing) + "\n")
            log(f"有 {len(missing)} 个编码在所选的场 / 状态里都没有查到（见「未找到的国标编码.txt」），"
                f"如 {missing[0]}；可能在其他场，请把场加进场列表后再查这些编码")

        if not pieces:
            write_record(task, 0, 0, queries, folder)
            log("这些国标编码在所选的场里都没有查到系谱，未生成合并文件")
            return {"path": folder, "total": 0, "rows": 0, "parts": 0}
        found = len(where)
        xlsx = [p for p in pieces if p["path"].lower().endswith(".xlsx")]
        if len(xlsx) == len(pieces):
            log(f"合并 {len(pieces)} 个文件 ...")
            tmp_out = os.path.join(folder, ".合并中.xlsx")
            rows = merge_with_source(pieces, tmp_out)
            final = move_replace(tmp_out, folder,
                                 sanitize(f"{stem}_找到{found}个_{gn._count_text(rows)}_合并"), ".xlsx")
        else:
            rows = None
            final = folder
            log("提示：导出的不是 xlsx 文件，无法合并，各次导出的文件都在任务文件夹里")
        write_record(task, found, rows, queries, final)
        log(f"[完成] {task_name(task)}：找到 {found}/{len(codes)} 个，查询 {queries} 次，"
            f"合并 {rows if rows is not None else '?'} 条 -> {folder}（用时 {fmt_secs(time.time() - t0)}）")
        return {"path": final, "total": rows, "rows": rows, "parts": len(pieces), "found": found,
                "missing": len(missing)}
    except BaseException:
        if not os.listdir(folder):
            shutil.rmtree(folder, ignore_errors=True)
        raise


# =====================================================================
# 会话 / 主流程
# =====================================================================
def _goto(driver):
    _MAIN["handle"] = None
    goto_page(driver)


def reload_page(driver):
    core.close_other_windows(driver, _MAIN.get("handle"))
    open_url(driver, CONFIG["PED_URL"])
    if not wait_until(lambda: on_page(driver), 20):
        core.back_to_home(driver)
        goto_page(driver)
    net_hook(driver)
    _MAIN["handle"] = _current_handle(driver)


ADAPTER = core.Adapter(goto=_goto, run=download_task, reload=reload_page, name=task_name,
                       download_dir=lambda: CONFIG["DOWNLOAD_DIR"], title=TITLE)


def run(tasks, username=None, password=None, status_fn=None):
    return core.run_tasks(tasks, ADAPTER, username, password, status_fn, what="任务")


def main():
    if "--headless" in sys.argv:
        core.HEADLESS = True

    def arg(name):
        return sys.argv[sys.argv.index(name) + 1] if name in sys.argv and sys.argv.index(name) + 1 < len(sys.argv) \
            else None
    if not arg("--codes") or not arg("--farms"):
        sys.exit('用法：python task_pedigree.py --codes 国标编码.txt --farms "龙马一场;龙马二场" '
                 '[--status 在场|离场|全部] [--headless]')
    st = {"在场": "只查在场", "离场": "只查离场"}.get(arg("--status") or "", STATUS_DEFAULT)
    task = make_task(gn.codes_from_file(arg("--codes")), re.split(r"[;；,，\n]+", arg("--farms")), st)
    username, password = core.load_credentials()
    run([task], username, password)


if __name__ == "__main__":
    main()
