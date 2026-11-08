# -*- coding: utf-8 -*-
"""
母猪档案 —— 按「场 × 阶段」批量下载
================================================================
首页「养殖生产」→「个体养殖」→「母猪养殖 / 母猪档案」→ 每个任务：
右上角选场 →「后备/生产阶段」或「仔猪/保育/待售阶段」→ 读取「共 N 条」→「导入导出 → 批量下载数据」
下载后：保存到「2026年 / 10月」（按下载日期），文件名 母猪档案_{场}_{阶段}_{N}条_{时间}.xlsx；表格最右侧追加「场」「阶段类型」列；
新增「下载信息」工作表；下载目录追加一行到「母猪档案下载记录.csv」。

    python task_sowfile.py                                  # 按 CONFIG["TASKS"] 下载
    python task_sowfile.py --tasks "龙马一场,1;龙马一场,2"   # 1=后备/生产阶段 2=仔猪/保育/待售阶段
    python task_sowfile.py --headless
"""

import os
import re
import sys
import csv
import time
import datetime
from copy import copy
from collections import ChainMap

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import InvalidSessionIdException, WebDriverException

import yz_core as core
from yz_core import (StopRequested, log, wait_until, fmt_secs, sanitize, safe_click, _js, _raise_if_dead,
                     _handles, _current_handle, _n, JS_VIS, JS_CLICK_BY_TEXT, JS_VISIBLE_TEXT, net_hook,
                     net_reset, net_state, net_dump, wait_loading_done, click_text, find_menu,
                     switch_to_new_window, save_debug, page_snapshot, open_url, table_state, wait_query_done,
                     query_total, is_limit_msg, watch_dirs, list_files, recent_partials, move_unique, check_stop)

TITLE = "母猪档案"

CONFIG = ChainMap({
    "DOWNLOAD_DIR": r"D:\test\downloads\母猪档案",
    "SOW_FILE_URL": "https://bpm.yingzi.com/#/person/sow/file/list",   # 菜单点不到时直接打开
    "DOWNLOAD_TIMEOUT": 600,
    "QUERY_TIMEOUT": 360,         # 切换场 / 阶段后等待数据刷新的最大秒数
    "QUERY_IDLE_GRACE": 12,
    "EXPORT_RETRY_WAIT": 120,     # 点导出后多少秒没文件就重试
    "EXPORT_NO_FILE_WAIT": 25,
    "EXPORT_ROW_LIMIT": 50000,
    "NO_REACTION_WAIT": 8,        # 点「批量下载数据」后几秒内没任何反应，就换一种点击方式
    "FARMS": ["龙马一场"],         # 界面版默认场列表
    "TASKS": [
        {"farm": "龙马一场", "stage": "后备/生产阶段"},
        {"farm": "龙马一场", "stage": "仔猪/保育/待售阶段"},
    ],
    "FILE_PREFIX": TITLE,
    "ADD_INFO_COLUMNS": True,     # 表格最右侧追加「场」「阶段类型」列
    "ADD_INFO_SHEET": True,       # 新增「下载信息」工作表
    "WRITE_RECORD": True,         # 追加到「母猪档案下载记录.csv」
    "SKIP_EMPTY": True,           # 页面显示 0 条时不导出，只记录数量
}, core.CONFIG)

STAGES = [
    {"name": "后备/生产阶段", "key": "stage1", "short": "后备生产阶段",
     "alias": ["后备", "生产", "后备阶段", "生产阶段", "1", "stage1"]},
    {"name": "仔猪/保育/待售阶段", "key": "stage2", "short": "仔猪保育待售阶段",
     "alias": ["活仔/保育/待售阶段", "仔猪", "活仔", "保育", "待售", "2", "stage2"]},
]
STAGE_NAMES = [s["name"] for s in STAGES]
DOWNLOAD_EXTS = (".xlsx", ".xls", ".csv", ".zip")

_MAIN = {"handle": None}


def resolve_stage(s):
    t = _n(s)
    for st in STAGES:
        if t == _n(st["name"]) or t in st["alias"]:
            return st
    for st in STAGES:
        if any(a in t for a in st["alias"] if len(a) > 1):
            return st
    raise ValueError(f"无法识别的阶段：{s}（可选：{' / '.join(STAGE_NAMES)}，或 1 / 2）")


def task_name(task):
    return f"{task['farm'].strip()}_{resolve_stage(task['stage'])['short']}"


# =====================================================================
# 页面内 JS
# =====================================================================
JS_STAGE_TAB = JS_VIS + """
const key = arguments[0], text = arguments[1];
for (const t of document.querySelectorAll('.ant-tabs-tab')) {
  if (!vis(t)) continue;
  const btn = t.querySelector('.ant-tabs-tab-btn') || t;
  if ((t.getAttribute('data-node-key') || '') === key || norm(btn.textContent) === text)
    return {el: btn, active: t.classList.contains('ant-tabs-tab-active')};
}
return null;
"""

JS_IO_TRIGGER = JS_VIS + """
const t = arguments[0];
for (const e of document.querySelectorAll('.ant-dropdown-trigger, button, a, span')) {
  if (vis(e) && norm(e.textContent) === t) return e.closest('.ant-dropdown-trigger') || e;
}
return null;
"""

# 页面结构是 li > span > button，下载动作绑在 button 上：必须点 button
JS_DROPDOWN_ITEM = JS_VIS + """
const t = arguments[0];
for (const d of document.querySelectorAll('.ant-dropdown:not(.ant-dropdown-hidden), .ant-dropdown-menu')) {
  if (!vis(d)) continue;
  for (const sel of ['button', 'a', '.ant-dropdown-menu-title-content', 'li', 'span'])
    for (const e of d.querySelectorAll(sel))
      if (vis(e) && norm(e.textContent) === t) return e.querySelector('button') || e;
}
return null;
"""

JS_EXPORT_MODAL = JS_VIS + """
for (const m of document.querySelectorAll('.ant-modal-wrap, .ant-modal-confirm, .el-dialog__wrapper, .el-message-box__wrapper, [role=dialog]')) {
  if (!vis(m) || m.closest('[class*=live-chat]')) continue;
  const txt = (m.innerText || '').trim();
  if (!txt) continue;
  const btns = Array.from(m.querySelectorAll('button')).filter(vis);
  const ok = btns.find(b => b.classList.contains('ant-btn-primary'))
          || btns.find(b => /确定|确认|导出|下载|是/.test(norm(b.textContent)));
  return {text: txt.replace(/\\s+/g, ' ').slice(0, 160), ok: ok || null};
}
return null;
"""

INFO_MSG_CSS = ".ant-message-notice, .ant-notification-notice"


def _escape(driver):
    try:
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    except Exception:
        pass


# =====================================================================
# 导航：个体养殖 → 母猪档案
# =====================================================================
def _stage_tab(driver, st):
    return _js(driver, JS_STAGE_TAB, st["key"], _n(st["name"]))


def on_sow_file(driver):
    return bool(_stage_tab(driver, STAGES[0]) or _stage_tab(driver, STAGES[1]))


def _click_menu(driver, title, timeout=10):
    el = wait_until(lambda: find_menu(driver, title), timeout)
    if el is None:
        return False
    before = set(_handles(driver))
    safe_click(driver, el)
    switch_to_new_window(driver, before, wait=1.0)
    return True


def goto_sow_file(driver, max_attempts=3):
    for attempt in range(1, max_attempts + 1):
        try:
            try:
                driver.switch_to.default_content()
            except WebDriverException as e:
                _raise_if_dead(e)
            if not on_sow_file(driver):
                if "bpm.yingzi.com" not in driver.current_url and not find_menu(driver, "个体养殖"):
                    log("进入「养殖生产」模块 ...")
                    before = set(_handles(driver))
                    try:
                        click_text(driver, "养殖生产", timeout=15)
                        switch_to_new_window(driver, before)
                        wait_loading_done(driver, timeout=60)
                        wait_until(lambda: find_menu(driver, "个体养殖"), CONFIG["TIMEOUT"])
                    except (StopRequested, InvalidSessionIdException):
                        raise
                    except Exception as e:
                        _raise_if_dead(e)
                        log(f"未能通过「养殖生产」进入（{str(e)[:60]}），直接打开母猪档案页面")
                        open_url(driver, CONFIG["SOW_FILE_URL"])

            if not on_sow_file(driver):
                log("进入「个体养殖」...")
                if _click_menu(driver, "个体养殖", 10):
                    wait_loading_done(driver, timeout=30)
                log("打开「母猪档案」...")
                item = wait_until(lambda: find_menu(driver, "母猪档案"), 5)
                if item is None and _click_menu(driver, "母猪养殖", 3):   # 左侧分组折叠着，先展开
                    item = wait_until(lambda: find_menu(driver, "母猪档案"), 5)
                if item is not None:
                    safe_click(driver, item)
                if not wait_until(lambda: on_sow_file(driver), 15):
                    log("菜单点击后未看到母猪档案页面，直接打开页面地址")
                    open_url(driver, CONFIG["SOW_FILE_URL"])
                    if "user/login" in driver.current_url:
                        raise RuntimeError("打开母猪档案页面时跳回了登录页（登录已失效）")

            if wait_until(lambda: on_sow_file(driver), CONFIG["TIMEOUT"]):
                wait_loading_done(driver, timeout=30)
                net_hook(driver)
                _MAIN["handle"] = _current_handle(driver)
                log("已进入「母猪档案」页面")
                return True
            log("未能进入母猪档案页面，准备重试 ...")
        except (StopRequested, InvalidSessionIdException):
            raise
        except Exception as e:
            _raise_if_dead(e)
            log(f"导航第 {attempt} 次尝试异常：{type(e).__name__}: {str(e)[:100]}")
        time.sleep(2)
    save_debug(driver, "进入母猪档案失败")
    raise RuntimeError("多次尝试后仍未能进入「母猪档案」页面，请检查网络或页面是否有变化。")


def ensure_sow_file(driver):
    mh = _MAIN.get("handle")
    if mh and _current_handle(driver) != mh and mh in _handles(driver):
        driver.switch_to.window(mh)
    if not on_sow_file(driver):
        goto_sow_file(driver)


# =====================================================================
# 场 / 阶段
# =====================================================================
def current_farm(driver):
    return core.current_org(driver)


def select_farm(driver, farm):
    """右上角场选择框选中指定场；切换后等列表数据刷新。返回是否发生了切换。"""
    before = table_state(driver)
    if not core.select_org(driver, farm):
        log(f"当前已是「{farm}」")
        return False
    wait_query_done(driver, before, time.time(), CONFIG["QUERY_TIMEOUT"], CONFIG["QUERY_IDLE_GRACE"], what="切换场")
    return True


def select_stage(driver, st):
    info = wait_until(lambda: _stage_tab(driver, st), 15)
    if not info:
        save_debug(driver, "未找到阶段标签")
        raise RuntimeError(f"页面上未找到「{st['name']}」标签")
    if info.get("active"):
        log(f"当前已是「{st['name']}」")
        return False
    log(f"切换到「{st['name']}」")
    net_hook(driver)
    net_reset(driver)
    before = table_state(driver)
    safe_click(driver, info["el"])
    if not wait_until(lambda: (_stage_tab(driver, st) or {}).get("active"), 5, poll=0.2):
        raise RuntimeError(f"点击「{st['name']}」后标签未切换")
    wait_query_done(driver, before, time.time(), CONFIG["QUERY_TIMEOUT"], CONFIG["QUERY_IDLE_GRACE"], what="切换阶段")
    return True


# =====================================================================
# 导出 / 下载
# =====================================================================
def _dirs():
    return watch_dirs(CONFIG["DOWNLOAD_DIR"])


def list_downloads():
    return list_files(_dirs(), DOWNLOAD_EXTS)


def click_batch_download(driver, mode=0):
    """「导入导出」→「批量下载数据」。mode：0=JS 点按钮，1=模拟真实鼠标点击，2=派发完整鼠标事件。"""
    for attempt in range(3):
        trig = wait_until(lambda: _js(driver, JS_IO_TRIGGER, "导入导出"), 15)
        if trig is None:
            return False
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", trig)
            ActionChains(driver).move_to_element(trig).perform()      # antd 下拉默认悬停展开
        except WebDriverException as e:
            _raise_if_dead(e)
        item = wait_until(lambda: _js(driver, JS_DROPDOWN_ITEM, "批量下载数据"), 2, poll=0.2)
        if item is None:
            safe_click(driver, trig)
            item = wait_until(lambda: _js(driver, JS_DROPDOWN_ITEM, "批量下载数据"), 3, poll=0.2)
        if item is not None:
            how = ["JS 点击", "鼠标点击", "鼠标事件"][mode % 3]
            try:
                if mode % 3 == 1:
                    ActionChains(driver).move_to_element(item).pause(0.2).click(item).perform()
                elif mode % 3 == 2:
                    driver.execute_script(
                        "for (const t of ['pointerdown','mousedown','pointerup','mouseup','click'])"
                        " arguments[0].dispatchEvent(new MouseEvent(t, {bubbles: true, cancelable: true, view: window}));",
                        item)
                else:
                    driver.execute_script("arguments[0].click();", item)
            except WebDriverException as e:
                _raise_if_dead(e)
                driver.execute_script("arguments[0].click();", item)
            log(f"已点击「导入导出 → 批量下载数据」（{how}）")
            return True
        log(f"下拉菜单未展开，重试（第 {attempt + 2} 次）")
        _escape(driver)
        time.sleep(0.5)
    return False


def _is_chrome(driver):
    try:
        return (driver.capabilities.get("browserName") or "").lower() in ("chrome", "chrome-headless-shell")
    except Exception:
        return False


JS_CHROME_DOWNLOADS = """
const m = document.querySelector('downloads-manager'); if (!m || !m.shadowRoot) return null;
const list = m.shadowRoot.querySelector('#downloadsList');
const items = (list && list.items) || [];
const s = v => (v && typeof v === 'object') ? (v.url || v.path || JSON.stringify(v)) : (v || '');
return items.slice(0, 8).map(i => ({name: s(i.fileName), path: s(i.filePath), state: String(i.state || ''),
  url: s(i.url).slice(0, 200), danger: String(i.dangerType || ''), reason: s(i.lastReasonText),
  started: Number(i.started || 0)}));
"""
SAFE_DANGER = ("", "0", "NOT_DANGEROUS", "USER_VALIDATED", "ACCEPTED", "ALLOWLISTED_BY_POLICY",
               "DEEP_SCANNED_SAFE", "UNDEFINED")


def chrome_downloads(driver):
    """打开 chrome://downloads 读取浏览器的下载记录（文件被拦截、下载失败、存到别的目录时靠它发现）。"""
    if not _is_chrome(driver):
        return None
    main = _current_handle(driver)
    items = None
    try:
        driver.switch_to.new_window("tab")
        driver.get("chrome://downloads")
        items = wait_until(lambda: _js(driver, JS_CHROME_DOWNLOADS), 3, poll=0.3)
    except (InvalidSessionIdException, StopRequested):
        raise
    except Exception:
        items = None
    finally:
        try:
            if _current_handle(driver) != main:
                driver.close()
        except WebDriverException:
            pass
        if main and main in _handles(driver):
            driver.switch_to.window(main)
    return items


def wait_export_file(driver, before, t_click, timeout, total=None, n_handles=None):
    """点「批量下载数据」后等待文件。返回 (path, reason)；path 为 None 时 reason：
       noclick = 点击后毫无反应；nofile = 请求已返回但没收到文件；timeout = 超时"""
    t0 = time.time()
    end = t0 + timeout
    seen_msgs, seen_modal, logged_req = set(), set(), False
    opened_urls = set()
    quiet_checked = False
    if n_handles is None:
        n_handles = len(_handles(driver))
    last_log = last_cd = t0
    net_quiet_since = None
    while time.time() < end:
        check_stop()
        now = time.time()

        new = list_downloads() - before
        if new:
            return max(new, key=os.path.getmtime), "ok"
        if recent_partials(_dirs()):
            end = max(end, now + 30)
            net_quiet_since = None
            if now - last_log >= 10:
                log("文件正在下载 ...")
                last_log = now
            time.sleep(0.5)
            continue

        md = _js(driver, JS_EXPORT_MODAL)
        if md and now - t0 < 60:
            key = md.get("text", "")
            if key not in seen_modal:
                seen_modal.add(key)
                log(f"导出弹框：{key}")
                if md.get("ok") is not None:
                    try:
                        driver.execute_script("arguments[0].click();", md["ok"])
                        log("已点击弹框中的确认按钮")
                    except WebDriverException as e:
                        _raise_if_dead(e)

        m = _js(driver, JS_VISIBLE_TEXT, INFO_MSG_CSS, default="") or ""
        if m and m not in seen_msgs:
            seen_msgs.add(m)
            log(f"页面提示：{m}")
            if is_limit_msg(m) and (total is None or total > CONFIG["EXPORT_ROW_LIMIT"]):
                raise RuntimeError(f"导出被平台拒绝：{m}")
            if re.search(r"失败|错误|异常|无权限|没有权限", m):
                raise RuntimeError(f"导出失败，页面提示：{m}")
            if re.search(r"下载中心|任务列表|后台|稍后|生成中", m):
                log("提示：平台是异步生成文件，继续等待；若一直没有文件，请把这条提示发给维护人员")
                end = max(end, now + CONFIG["DOWNLOAD_TIMEOUT"])

        st = net_state(driver) or {}
        dump = net_dump(driver)
        if st.get("s", 0) > 0 and not logged_req and st.get("p", 0) == 0:
            logged_req = True
            for r in (dump.get("last") or [])[-4:]:
                log(f"  导出请求：{r.get('u')} → {r.get('s')}（{r.get('ms')} ms，{r.get('ct') or '-'}）")
                if r.get("s") and r["s"] >= 400:
                    log(f"警告：导出接口返回状态 {r['s']}")

        # 接口返回里带了文件地址，但浏览器没开始下载：直接打开该地址下载
        for u in dump.get("urls") or []:
            if u not in opened_urls and now - t0 > 4:
                opened_urls.add(u)
                log(f"接口返回了文件地址，直接打开下载：{u[:120]}")
                try:
                    driver.execute_script("window.open(arguments[0], '_blank');", u)
                except WebDriverException as e:
                    _raise_if_dead(e)
                end = max(end, now + 120)

        if (st.get("s", 0) == 0 and not seen_modal and not seen_msgs and not quiet_checked
                and now - t0 >= CONFIG.get("NO_REACTION_WAIT", 8)):
            quiet_checked = True
            started = [it for it in (chrome_downloads(driver) or [])
                       if it.get("started", 0) / 1000.0 >= t_click - 5]
            if len(_handles(driver)) == n_handles and not started:
                return None, "noclick"
            log("页面没有发接口请求，但浏览器已开始下载 / 打开了新窗口，继续等待 ...")
            last_cd = now

        if st.get("s", 0) > 0 and st.get("p", 0) == 0:
            net_quiet_since = net_quiet_since or now
        else:
            net_quiet_since = None

        if now - last_cd >= 15:
            last_cd = now
            for it in chrome_downloads(driver) or []:
                if it.get("started") and it["started"] / 1000.0 < t_click - 5:
                    continue
                stt, pth = it.get("state", "").upper(), it.get("path", "")
                if stt in ("COMPLETE", "2") and pth and os.path.isfile(pth) and os.path.getmtime(pth) >= t_click - 5:
                    log(f"浏览器已下载到：{pth}")
                    return pth, "ok"
                if stt in ("INTERRUPTED", "DANGEROUS", "INSECURE", "3", "4") or \
                        it.get("danger", "").upper() not in SAFE_DANGER:
                    raise RuntimeError(f"浏览器拦截/中断了下载：{it.get('name')} 状态 {stt} "
                                       f"{it.get('reason') or it.get('danger')}")
                if stt in ("IN_PROGRESS", "0"):
                    end = max(end, now + 30)
                    net_quiet_since = None

        if net_quiet_since and now - net_quiet_since >= CONFIG["EXPORT_NO_FILE_WAIT"] and not opened_urls:
            return None, "nofile"
        if now - last_log >= 20:
            log(f"等待下载文件 ... 已等待 {fmt_secs(now - t0)}（请求 {st.get('s', 0)} 个，未完成 {st.get('p', 0)} 个）")
            last_log = now
        time.sleep(0.5)
    return None, "timeout"


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


def annotate_xlsx(path, farm, st, total):
    """表格最右侧追加「场」「阶段类型」列，新增「下载信息」工作表。返回文件数据行数（失败为 None）。"""
    if not (CONFIG.get("ADD_INFO_COLUMNS") or CONFIG.get("ADD_INFO_SHEET")):
        return None
    if not path.lower().endswith(".xlsx"):
        log(f"提示：下载的是 {os.path.splitext(path)[1]} 文件，只能写入文件名，不能在表格里追加场/阶段列")
        return None
    try:
        import openpyxl
    except ImportError:
        log("提示：未安装 openpyxl，跳过在表格内追加信息")
        return None
    try:
        wb = openpyxl.load_workbook(path)
        ws = wb.worksheets[0]
        header_end = 1
        for mr in ws.merged_cells.ranges:          # 多级表头：第 1 行开始的合并单元格
            if mr.min_row == 1:
                header_end = max(header_end, mr.max_row)
        max_col = ws.max_column
        data_rows = [idx for idx, vals in enumerate(
            ws.iter_rows(min_row=header_end + 1, max_col=max_col, values_only=True), start=header_end + 1)
            if any(v not in (None, "") for v in vals)]
        rows = len(data_rows)

        if CONFIG.get("ADD_INFO_COLUMNS"):
            src = ws.cell(1, max_col)
            for off, title in ((1, "场"), (2, "阶段类型")):
                c = ws.cell(1, max_col + off, title)
                if src.has_style:
                    c.font, c.fill, c.border, c.alignment = (copy(src.font), copy(src.fill),
                                                             copy(src.border), copy(src.alignment))
                if header_end > 1:
                    ws.merge_cells(start_row=1, start_column=max_col + off,
                                   end_row=header_end, end_column=max_col + off)
                ws.column_dimensions[c.column_letter].width = 16 if off == 1 else 20
            for r in data_rows:
                ws.cell(r, max_col + 1, farm)
                ws.cell(r, max_col + 2, st["name"])

        if CONFIG.get("ADD_INFO_SHEET"):
            name = "下载信息"
            if name in wb.sheetnames:
                del wb[name]
            info = wb.create_sheet(name)
            for row in (("项目", "内容"), ("场", farm), ("阶段类型", st["name"]),
                        ("页面显示条数", total if total is not None else "未读取到"),
                        ("文件数据行数", rows),
                        ("下载时间", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))):
                info.append(row)
            info.column_dimensions["A"].width = 16
            info.column_dimensions["B"].width = 28
        wb.save(path)
        return rows
    except Exception as e:
        log(f"警告：向表格追加信息失败（{type(e).__name__}: {str(e)[:80]}），文件保持原样")
        return None


def write_record(farm, st, total, rows, path):
    if not CONFIG.get("WRITE_RECORD"):
        return
    rec = os.path.join(CONFIG["DOWNLOAD_DIR"], f"{CONFIG['FILE_PREFIX']}下载记录.csv")
    new = not os.path.exists(rec)
    try:
        with open(rec, "a", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            if new:
                w.writerow(["下载时间", "场", "阶段类型", "页面显示条数", "文件数据行数", "文件"])
            w.writerow([datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), farm, st["name"],
                        "" if total is None else total, "" if rows is None else rows,
                        core.rel_to(path, CONFIG["DOWNLOAD_DIR"]) if path else "（无数据，未导出）"])
    except OSError as e:
        log(f"警告：写入下载记录失败：{e}（若记录表正被 Excel 打开，请关闭后再试）")


def download_task(driver, task):
    """下载一个「场 × 阶段」任务。返回 {"path", "total", "rows"}。"""
    farm = task["farm"].strip()
    st = resolve_stage(task["stage"])
    t0 = time.time()
    log(f"---- 开始：{farm} · {st['name']} ----")

    ensure_sow_file(driver)
    _escape(driver)
    select_farm(driver, farm)
    select_stage(driver, st)
    if not (_stage_tab(driver, st) or {}).get("active"):      # 切换场后页面可能重置标签
        select_stage(driver, st)
    if _n(current_farm(driver)) != _n(farm):
        raise RuntimeError(f"场选择框当前为「{current_farm(driver)}」，不是「{farm}」")
    t_set = time.time()

    total = query_total(driver)
    log(f"{farm} · {st['name']}：共 {total if total is not None else '（未读取到）'} 条")
    if total == 0 and CONFIG.get("SKIP_EMPTY"):
        log("没有数据，跳过导出")
        write_record(farm, st, 0, 0, "")
        return None                        # 界面显示「无数据」

    before = list_downloads()
    handles_before = set(_handles(driver))
    path, reason = None, ""
    for attempt in range(3):
        net_hook(driver)
        net_reset(driver)
        t_click = time.time()
        n_before = len(_handles(driver))
        if not click_batch_download(driver, mode=attempt):
            save_debug(driver, f"未找到批量下载_{farm}")
            raise RuntimeError("未能点开「导入导出 → 批量下载数据」")
        path, reason = wait_export_file(driver, before, t_click, CONFIG["EXPORT_RETRY_WAIT"], total, n_before)
        if path:
            break
        _close_new_windows(driver, handles_before)
        _escape(driver)
        if reason == "noclick":
            log("点击后页面没有任何反应（没有发出请求、没有弹框），换一种点击方式重试 ...")
        elif reason == "nofile":
            log("导出请求已返回，但没有收到文件，重试导出 ...")
        else:
            log("等待超时，重试导出 ...")
    if not path:
        for r in net_dump(driver).get("last") or []:
            log(f"  最近请求：{r.get('u')} → {r.get('s')} {r.get('ct') or ''} {r.get('cd') or ''}")
        for it in (chrome_downloads(driver) or [])[:3]:
            log(f"  浏览器下载记录：{it.get('name')} 状态 {it.get('state')} {it.get('reason') or ''}")
        page_snapshot(driver)
        save_debug(driver, f"下载失败_{farm}")
        hint = {"noclick": "三种方式点击「批量下载数据」都没有反应",
                "nofile": "导出请求返回了，但浏览器没有下载到文件",
                "timeout": "等待超时"}.get(reason, reason)
        raise RuntimeError(f"{hint}（{farm} · {st['name']}）。上面的「最近请求」和 debug 文件夹里的截图可用于排查")
    t_dl = time.time()

    file_rows = core.xlsx_data_rows(path)          # 文件名里的条数以导出文件为准（页面上的数字可能读错）
    if file_rows is not None and total is not None and file_rows != total:
        log(f"提示：页面显示 {total} 条，导出文件实际 {file_rows} 条，以文件为准")
    n = file_rows if file_rows is not None else total
    cnt = f"{n}条" if n is not None else "条数未知"
    stem = sanitize(f"{CONFIG['FILE_PREFIX']}_{farm}_{st['short']}_{cnt}_{time.strftime('%Y%m%d_%H%M')}")
    today = datetime.date.today()          # 母猪档案是当天的快照：按下载日期放进「年 / 月」文件夹
    final = move_unique(path, core.period_folder(CONFIG["DOWNLOAD_DIR"], today, today), stem)
    _close_new_windows(driver, handles_before)
    rows = annotate_xlsx(final, farm, st, total)
    if rows is not None and total is not None and rows != total:
        log(f"警告：文件数据行数 {rows} 与页面显示 {total} 条不一致，请核对")
    write_record(farm, st, total, rows, final)
    log(f"[完成] {farm} · {st['name']} -> {final}（{os.path.getsize(final) / 1024:.0f} KB，"
        f"{rows if rows is not None else total} 条）")
    log(f"耗时明细：设置条件 {fmt_secs(t_set - t0)} · 导出下载 {fmt_secs(t_dl - t_set)} · "
        f"合计 {fmt_secs(time.time() - t0)}")
    return {"path": final, "total": n, "rows": rows if rows is not None else file_rows}


# =====================================================================
# 会话 / 主流程
# =====================================================================
def _goto(driver):
    _MAIN["handle"] = None
    goto_sow_file(driver)


def open_session(username, password):
    return core.start_session(CONFIG["DOWNLOAD_DIR"], username, password, _goto)


def reload_page(driver):
    """不重新登录，直接重新打开母猪档案页面（清掉残留弹窗和多余窗口）。"""
    core.close_other_windows(driver, _MAIN.get("handle"))
    open_url(driver, CONFIG["SOW_FILE_URL"])
    if "user/login" in driver.current_url:
        raise RuntimeError("登录已失效")
    if not wait_until(lambda: on_sow_file(driver), 20):
        core.back_to_home(driver)
        goto_sow_file(driver)
    net_hook(driver)
    _MAIN["handle"] = _current_handle(driver)


ADAPTER = core.Adapter(goto=_goto, run=download_task, reload=reload_page, name=task_name,
                       download_dir=lambda: CONFIG["DOWNLOAD_DIR"], title=TITLE)


def run(tasks, username=None, password=None, status_fn=None):
    for t in tasks:
        resolve_stage(t["stage"])          # 提前校验阶段写法
    return core.run_tasks(tasks, ADAPTER, username, password, status_fn, what="任务")


def parse_tasks(spec):
    """'龙马一场,1;龙马二场,仔猪/保育/待售阶段' → 任务列表"""
    out = []
    for item in spec.split(";"):
        item = item.strip()
        if not item:
            continue
        parts = [p.strip() for p in re.split(r"[,，]", item, maxsplit=1)]
        if len(parts) != 2 or not parts[0]:
            raise ValueError(f"任务格式应为「场名,阶段」：{item}")
        out.append({"farm": parts[0], "stage": resolve_stage(parts[1])["name"]})
    return out


def main():
    if "--headless" in sys.argv:
        core.HEADLESS = True
    tasks = CONFIG["TASKS"]
    if "--tasks" in sys.argv:
        i = sys.argv.index("--tasks")
        if i + 1 < len(sys.argv):
            tasks = parse_tasks(sys.argv[i + 1])
    username, password = core.load_credentials()
    if tasks:
        run(tasks, username, password)


if __name__ == "__main__":
    main()
