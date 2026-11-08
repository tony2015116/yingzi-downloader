# -*- coding: utf-8 -*-
"""
育种值查询（基因组选种平台）—— 按国标编码批量导出
================================================================
首页「基因组选种」→ 左侧「育种值」→「育种值查询」→ 标签「按个体」→ 每批：
点「批量」粘贴国标编码（一次最多 400 个）→ 查询 → 右上角「导出」。
编码超过 400 个时每 400 个一批分别导出，最后合并成一个文件。只用国标编码这一个条件，不填日期等其他条件。

文件：育种值查询\\2026年\\10月\\育种值_按个体_国标N个_{N}条_{年月日_时分秒}.xlsx
（育种值会随计算更新，每次下载都是当时的结果，所以文件名带下载时间，不覆盖以前下载的文件）
每个任务追加一行到「育种值查询下载记录.csv」。

    python task_ebv.py --codes 国标编码.txt [--headless]
"""

import os
import re
import sys
import time
import shutil
import datetime
from collections import ChainMap

from selenium.common.exceptions import InvalidSessionIdException, WebDriverException

import yz_core as core
import task_genetic as gn
from yz_core import (StopRequested, log, wait_until, fmt_secs, sanitize, safe_click, _js, _raise_if_dead,
                     _handles, _current_handle, net_hook, net_reset, net_dump, wait_loading_done, click_text,
                     find_menu, switch_to_new_window, save_debug, open_url, page_snapshot, watch_dirs, list_files,
                     move_replace, period_folder, rel_to, xlsx_data_rows, append_csv)

TITLE = "育种值查询"

CONFIG = ChainMap({
    "EBV_URL": "https://gsm.yingzi.com/#/genomicSelection/query",     # 菜单点不到时直接打开
    "DOWNLOAD_DIR": r"D:\test\downloads\育种值查询",
    "TAB": "按个体",
    "TAB_KEY": "1",
    "CODE_INPUT_ID": "qp-personId-in",
}, gn.CONFIG)          # 平台地址、等待时间、5 万条上限、每批 400 个等与遗传性状查询共用

FILE_PREFIX = "育种值"
RECORD_NAME = f"{TITLE}下载记录.csv"
_MAIN = {"handle": None}


# =====================================================================
# 任务 / 命名
# =====================================================================
def make_task(codes, tab=None):
    codes, _bad = gn.parse_codes("\n".join(codes or []))
    return {"tab": tab or CONFIG["TAB"], "codes": codes}


def task_name(task):
    c = task["codes"]
    return f"{task.get('tab') or CONFIG['TAB']} · 国标 {len(c)} 个" + (f"（{c[0]} 等）" if c else "")


def make_stem(task):
    return sanitize(f"{FILE_PREFIX}_{task.get('tab') or CONFIG['TAB']}_国标{len(task['codes'])}个")


# =====================================================================
# 导航：首页「基因组选种」→ 育种值 → 育种值查询 → 按个体
# =====================================================================
JS_ON_PAGE = core.JS_VIS + """
const title = Array.from(document.querySelectorAll('.ant-page-header-heading-title')).filter(vis).map(e => norm(e.textContent));
const tabs = Array.from(document.querySelectorAll('.ant-tabs-tab')).filter(vis).map(t => norm(t.textContent));
return title.includes('育种值查询') && tabs.includes(arguments[0]);
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
                        log(f"未能通过首页「{CONFIG['PORTAL_TILE']}」进入（{str(e)[:60]}），直接打开育种值查询页面")
                        open_url(driver, CONFIG["EBV_URL"])
                if CONFIG["GSM_HOST"] in gn._url(driver) and not on_page(driver):
                    item = wait_until(lambda: find_menu(driver, "育种值查询"), 3)
                    if item is None:
                        grp = wait_until(lambda: find_menu(driver, "育种值"), 8)
                        if grp is not None:
                            log("展开左侧「育种值」...")
                            safe_click(driver, grp)
                            item = wait_until(lambda: find_menu(driver, "育种值查询"), 5)
                    if item is not None:
                        log("打开「育种值查询」...")
                        safe_click(driver, item)
                    if not wait_until(lambda: on_page(driver), 15):
                        log("菜单点击后未看到育种值查询页面，直接打开页面地址")
                        open_url(driver, CONFIG["EBV_URL"])
            if gn._login_lost(driver):
                raise RuntimeError("打开基因组选种平台时跳回了登录页（登录已失效或没有该平台权限）")
            if wait_until(lambda: on_page(driver), CONFIG["TIMEOUT"]):
                wait_loading_done(driver, timeout=30)
                net_hook(driver)
                _MAIN["handle"] = _current_handle(driver)
                log("已进入「育种值查询」页面")
                return True
            log("未能进入育种值查询页面，准备重试 ...")
        except (StopRequested, InvalidSessionIdException):
            raise
        except Exception as e:
            _raise_if_dead(e)
            log(f"导航第 {attempt} 次尝试异常：{type(e).__name__}: {str(e)[:100]}")
        time.sleep(2)
    save_debug(driver, "进入育种值查询失败")
    raise RuntimeError("多次尝试后仍未能进入「育种值查询」页面，请确认账号有「基因组选种」平台的权限。")


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


JS_CODE_INPUT = core.JS_VIS + """
const e = document.getElementById(arguments[0]);
if (e && vis(e) && !e.closest('.ant-modal-wrap')) return true;
// id 对不上时：页面上出现了带「批量」的输入框也算加载好了
return Array.from(document.querySelectorAll('.ant-input-group-addon')).some(a => vis(a)
  && !a.closest('.ant-modal-wrap') && norm(a.textContent) === '批量');
"""


def wait_form(driver):
    """育种值查询的查询表单在标签下方单独的卡片里，进入页面后要等一会儿才渲染：等国标编码输入框出现。"""
    if wait_until(lambda: _js(driver, JS_CODE_INPUT, CONFIG["CODE_INPUT_ID"], default=False), 3, poll=0.3):
        return
    log("等待查询条件区域加载 ...")
    if wait_until(lambda: _js(driver, JS_CODE_INPUT, CONFIG["CODE_INPUT_ID"], default=False),
                  CONFIG["TIMEOUT"], poll=0.5):
        wait_loading_done(driver, timeout=30)
        return
    save_debug(driver, "育种值查询_未找到国标编码输入框")
    raise RuntimeError("育种值查询页面上一直没有出现「国标编码」输入框（debug 文件夹里有当时的截图）")


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
# 一批：批量粘贴编码 → 查询 → 导出
# =====================================================================
def _dirs():
    return watch_dirs(CONFIG["DOWNLOAD_DIR"])


def export_batch(driver, codes):
    """返回 (临时文件, 条数, 页面条数)；0 条返回 None。"""
    ensure_page(driver)
    gn._escape(driver)
    select_tab(driver)
    wait_form(driver)
    # 不点「重置」：每批点「批量」时弹窗里的编码会被清空后重新粘贴，确定后就替换掉上一批的编码
    gn.set_codes(driver, codes, CONFIG["CODE_INPUT_ID"])
    total = gn.run_query(driver)
    log(f"查询结果：共 {total if total is not None else '（未读取到）'} 条")
    if total == 0:
        return None
    if total is not None and total > CONFIG["EXPORT_ROW_LIMIT"]:
        raise RuntimeError(f"本批查询结果 {total} 条，超过单次导出上限 {CONFIG['EXPORT_ROW_LIMIT']} 条"
                           "（育种值查询只按国标编码查，无法再按日期拆分）")
    before = list_files(_dirs(), gn.DOWNLOAD_EXTS)
    handles_before = set(_handles(driver))
    t0 = time.time()
    path = None
    for attempt in range(3):
        btn, why = gn.export_ready(driver, total)
        if btn is None:
            save_debug(driver, "导出按钮不可用_育种值")
            raise RuntimeError(f"「导出」不可用：{why}")
        net_hook(driver)
        net_reset(driver)
        safe_click(driver, btn)
        log("已点击「导出」" + (f"（第 {attempt + 1} 次）" if attempt else ""))
        try:
            path = gn.wait_download(driver, before, CONFIG["EXPORT_RETRY_WAIT"], total, dirs=_dirs())
        except gn.NeedSplit as e:
            raise RuntimeError(f"{e}（育种值查询只按国标编码查，无法再拆分）")
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
        save_debug(driver, "下载超时_育种值")
        raise RuntimeError("等待下载文件超时，请检查导出是否触发（debug 文件夹里有当时的截图）")
    _close_new_windows(driver, handles_before)
    rows = xlsx_data_rows(path)
    if rows == 0 and total:
        log(f"警告：导出文件读出 0 条，但页面显示 {total} 条，条数改用页面数字，请打开文件核对")
        rows = None
    if rows is not None and total is not None and rows != total:
        log(f"提示：页面显示 {total} 条，导出文件实际 {rows} 条，以文件为准")
    log(f"导出完成：{rows if rows is not None else total} 条（用时 {fmt_secs(time.time() - t0)}）")
    return path, (rows if rows is not None else total), total


def _unique_stamp(base, stem, count_text):
    """文件名里的下载时间：精确到秒；同一秒里已有同名文件（几乎不会）时再加 _2、_3，绝不覆盖以前的文件。"""
    stamp = time.strftime("%Y%m%d_%H%M%S")
    cand, k = stamp, 1
    while os.path.exists(os.path.join(base, f"{stem}_{count_text}_{cand}.xlsx")):
        k += 1
        cand = f"{stamp}_{k}"
    return cand


def write_record(task, page, rows, batches, path):
    append_csv(os.path.join(CONFIG["DOWNLOAD_DIR"], RECORD_NAME),
               ["下载时间", "标签", "国标编码个数", "编码批数", "页面显示条数（参考）", "文件数据行数", "保存位置"],
               [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), task.get("tab") or CONFIG["TAB"],
                len(task["codes"]), batches, "" if page is None else page, "" if rows is None else rows,
                rel_to(path, CONFIG["DOWNLOAD_DIR"]) if path else "（无数据，未导出）"])


def download_task(driver, task):
    codes = list(task.get("codes") or [])
    if not codes:
        raise ValueError("育种值查询必须填写国标编码")
    t0 = time.time()
    today = datetime.date.today()
    log(f"---- 开始：{task_name(task)} ----")
    base = period_folder(CONFIG["DOWNLOAD_DIR"], today, today)      # 没有日期条件：按下载日期放进年 / 月文件夹
    stem = make_stem(task)
    work = os.path.join(base, f".{stem}_临时_{time.strftime('%H%M%S')}")
    os.makedirs(work, exist_ok=True)
    batches = gn.code_batches(codes, CONFIG["CODE_BATCH"])
    if len(batches) > 1:
        log(f"国标编码 {len(codes)} 个，超过 {CONFIG['CODE_BATCH']} 个：分 {len(batches)} 批导出，最后合并成一个文件")
    pieces = []
    try:
        for k, batch in enumerate(batches, 1):
            if len(batches) > 1:
                log(f"== 第 {k}/{len(batches)} 批：{len(batch)} 个编码 ==")
                core.progress(k - 1, len(batches), "批")
            res = export_batch(driver, batch)
            if res is None:
                log("本批没有数据")
                continue
            tmp, n, page = res
            dst = move_replace(tmp, work, f"{k:03d}", os.path.splitext(tmp)[1].lower() or ".xlsx")
            pieces.append({"path": dst, "rows": n, "page": page})
        if len(batches) > 1:
            core.progress(len(batches), len(batches), "批")
        if not pieces:
            write_record(task, 0, 0, len(batches), "")
            log("这些国标编码都没有育种值数据，未生成文件")
            return None
        rows = sum(p["rows"] for p in pieces) if all(p["rows"] is not None for p in pieces) else None
        pages = [p["page"] for p in pieces]
        page = sum(pages) if all(x is not None for x in pages) else None
        stamp = _unique_stamp(base, stem, gn._count_text(rows))
        if len(pieces) == 1:
            p = pieces[0]
            final = move_replace(p["path"], base, f"{stem}_{gn._count_text(rows)}_{stamp}",
                                 os.path.splitext(p["path"])[1])
        elif all(p["path"].lower().endswith(".xlsx") for p in pieces):
            log(f"合并 {len(pieces)} 个文件 ...")
            tmp_out = os.path.join(work, "合并.xlsx")
            merged = gn.merge_xlsx([p["path"] for p in pieces], tmp_out)
            if rows is not None and merged != rows:
                log(f"警告：合并后 {merged} 条，各文件合计 {rows} 条，请核对")
            rows = merged
            final = move_replace(tmp_out, base, f"{stem}_{gn._count_text(rows)}_{stamp}", ".xlsx")
        else:
            seg = os.path.join(base, f"{stem}_{stamp}_分批")
            os.makedirs(seg, exist_ok=True)
            for k, p in enumerate(pieces, 1):
                move_replace(p["path"], seg, f"{stem}_第{k}批_{gn._count_text(p['rows'])}",
                             os.path.splitext(p["path"])[1])
            log("提示：导出的不是 xlsx 文件，无法合并，各批文件放在同一个文件夹里")
            final = seg
        write_record(task, page, rows, len(batches), final)
        size = f"{os.path.getsize(final) / 1024:.0f} KB，" if os.path.isfile(final) else ""
        log(f"[完成] {task_name(task)} -> {final}（{size}{rows if rows is not None else page} 条，"
            f"用时 {fmt_secs(time.time() - t0)}）")
        return {"path": final, "total": rows if rows is not None else page, "rows": rows, "parts": 1}
    finally:
        shutil.rmtree(work, ignore_errors=True)


# =====================================================================
# 会话 / 主流程
# =====================================================================
def _goto(driver):
    _MAIN["handle"] = None
    goto_page(driver)


def reload_page(driver):
    core.close_other_windows(driver, _MAIN.get("handle"))
    open_url(driver, CONFIG["EBV_URL"])
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
    if "--codes" not in sys.argv:
        sys.exit("用法：python task_ebv.py --codes 国标编码.txt [--headless]")
    i = sys.argv.index("--codes")
    task = make_task(gn.codes_from_file(sys.argv[i + 1]))
    username, password = core.load_credentials()
    run([task], username, password)


if __name__ == "__main__":
    main()
