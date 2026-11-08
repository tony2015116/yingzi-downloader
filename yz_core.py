# -*- coding: utf-8 -*-
"""
影子科技（扬翔）养殖平台下载器 —— 公共模块
================================================================
三个下载功能（母猪综合查询 / 猪只转移 / 母猪档案）共用的部分：
  运行开关（停止、后台运行、日志、进度）· 浏览器启动 · 登录（记住登录状态 / 验证码 / 自动重试）
  · 等待与点击工具 · 网络请求监控 · 下载目录监控 · 查询完成判断 · 通用任务循环
"""

import os
import re
import sys
import csv
import time
import shutil
import zipfile
import datetime

# 打包成无控制台的 exe 时 sys.stdout / sys.stderr 为 None，print 和第三方库写日志会出错
for _s in ("stdout", "stderr"):
    if getattr(sys, _s, None) is None:
        setattr(sys, _s, open(os.devnull, "w", encoding="utf-8"))
    try:
        getattr(sys, _s).reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# 显式导入 Chrome/Edge 模块：新版 selenium 是延迟导入，PyInstaller 打包时扫描不到
from selenium.webdriver.chrome.webdriver import WebDriver as ChromeDriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.webdriver import WebDriver as EdgeDriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import (
    TimeoutException,
    ElementClickInterceptedException,
    ElementNotInteractableException,
    NoSuchElementException,
    StaleElementReferenceException,
    InvalidSessionIdException,
    WebDriverException,
)

# =====================================================================
# 运行开关（界面版挂接；命令行模式不受影响）
# =====================================================================
LOG_HOOK = None        # callable(line)：日志显示到窗口
STOP_EVENT = None      # threading.Event：点「停止」时触发
HEADLESS = False       # True = 后台运行（不显示浏览器）
CURRENT_DRIVER = None  # 当前浏览器，退出 / 强制停止时关闭
PROGRESS_HOOK = None   # callable(n, total, unit)：任务内的细分进度（猪只转移逐条导出、母猪综合查询分段时用）
SESSION = {"home_url": None, "browser_dir": None, "current_dir": None}

CONFIG = {
    "BASE_URL": "https://account.yingzi.com/#/user/login",
    "CHROME_PATH": "",                 # 可选：手动指定 Chrome 路径（一般自动检测）
    "TIMEOUT": 30,                     # 一般步骤最多等多久（秒），条件满足立即继续
    "POLL": 0.4,
    "EXTRA_DOWNLOAD_DIRS": [os.path.expanduser("~/Downloads")],
    "DEBUG_DIR": "",                   # 出错截图目录（留空 = 程序目录下 debug）
    "KEEP_LOGIN": True,                # 记住登录状态
    "PROFILE_DIR": "",                 # 留空 = 程序目录下 browser_profile
    "LOGIN_ATTEMPTS": 3,
    "CAPTCHA_WAIT": 180,
    "DATE_FOLDERS": True,              # 按日期范围自动放进「2026年 / 06月」子文件夹
}


class StopRequested(BaseException):
    """用户点「停止」。继承 BaseException：各处的 `except Exception` 不会把它吞掉。"""


class LoginFatal(Exception):
    """账号密码错误、后台运行遇到验证码等 —— 重试也没用的登录失败。"""


# =====================================================================
# 基础工具
# =====================================================================
def _app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _first_existing(paths):
    for p in paths:
        if p and os.path.isfile(p):
            return p
    return None


def log(msg):
    line = f"[{time.strftime('%H:%M:%S')}] {msg}"
    try:
        print(line, flush=True)
    except Exception:
        pass
    if LOG_HOOK:
        try:
            LOG_HOOK(line)
        except Exception:
            pass


def progress(n, total, unit="条"):
    if PROGRESS_HOOK:
        try:
            PROGRESS_HOOK(n, total, unit)
        except Exception:
            pass


def _stopped():
    return STOP_EVENT is not None and STOP_EVENT.is_set()


def check_stop():
    if _stopped():
        raise StopRequested("用户已停止下载")


_DEAD_MARKERS = ("invalid session id", "chrome not reachable", "session deleted", "disconnected",
                 "no such session", "target closed", "connection refused", "max retries exceeded",
                 "failed to establish a new connection", "connection aborted", "connectionreseterror")


def is_dead_error(e):
    """浏览器已关闭 / 崩溃 / 被强制结束。"""
    if isinstance(e, InvalidSessionIdException):
        return True
    s = f"{type(e).__name__} {e}".lower()
    return any(k in s for k in _DEAD_MARKERS)


def _raise_if_dead(e):
    """浏览器已失效时统一转成 InvalidSessionIdException，交给上层自动重启。"""
    if isinstance(e, InvalidSessionIdException):
        raise e
    if is_dead_error(e):
        raise InvalidSessionIdException(str(e))


def wait_until(fn, timeout, poll=None):
    """轮询 fn()，返回第一个真值；超时返回 None。期间响应「停止」，忽略元素失效类异常。"""
    poll = poll or CONFIG["POLL"]
    end = time.time() + timeout
    while True:
        check_stop()
        try:
            v = fn()
            if v:
                return v
        except (StaleElementReferenceException, NoSuchElementException):
            pass
        except WebDriverException as e:
            _raise_if_dead(e)
        except Exception as e:
            if is_dead_error(e):
                raise InvalidSessionIdException(str(e))
            raise
        if time.time() >= end:
            return None
        time.sleep(poll)


def fmt_secs(sec):
    sec = int(sec)
    return f"{sec // 60}分{sec % 60:02d}秒" if sec >= 60 else f"{sec}秒"


def sanitize(name):
    return re.sub(r'[\\/:*?"<>|\s]+', "_", str(name)).strip("_") or "download"


def _n(s):
    return re.sub(r"\s+", "", str(s or ""))


def _digits(s):
    return re.sub(r"\D", "", str(s or ""))


def safe_click(driver, el):
    """滚动到元素并点击；被遮挡或失效时重试，最后用 JS 兜底。"""
    for _ in range(3):
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            el.click()
            return
        except (ElementClickInterceptedException, StaleElementReferenceException,
                ElementNotInteractableException):
            time.sleep(0.25)
        except WebDriverException as e:
            _raise_if_dead(e)
            time.sleep(0.25)
    try:
        driver.execute_script("arguments[0].click();", el)
    except WebDriverException as e:
        _raise_if_dead(e)


def _js(driver, script, *args, default=None):
    try:
        return driver.execute_script(script, *args)
    except WebDriverException as e:
        _raise_if_dead(e)
        return default


# =====================================================================
# 页面内 JS
# =====================================================================
JS_VIS = """
const vis = e => { if (!e) return false;
  if (!(e.offsetWidth || e.offsetHeight || e.getClientRects().length)) return false;
  const cs = getComputedStyle(e); return cs.visibility !== 'hidden' && cs.display !== 'none'; };
const norm = s => (s || '').replace(/\\s+/g, '');
"""

JS_CLICK_BY_TEXT = JS_VIS + """
const texts = arguments[0], scopeSel = arguments[1], elSel = arguments[2];
const roots = scopeSel ? Array.from(document.querySelectorAll(scopeSel)).filter(vis) : [document];
for (const t of texts) for (const root of roots)
  for (const e of root.querySelectorAll(elSel))
    if (norm(e.textContent) === t && vis(e)) { e.click(); return t; }
return null;
"""

# 只认明确的加载动画（[class*='loading'] 会匹配到常驻元素，导致一直误判「加载中」）
SPIN_CSS = (".ant-spin-spinning, .ant-spin-blur, .davinci-spin-spinning, .davinci-spin-blur, .el-loading-mask, "
            "[class*='loading-mask'], [class*='spinner']")     # davinci-*：母猪综合查询报表（DaVinci）自己的加载动画
JS_SPIN_IN = JS_VIS + """
const root = arguments[0] || document;
return Array.from(root.querySelectorAll(arguments[1])).some(vis);
"""

JS_SET_VALUE = """
const el = arguments[0], v = arguments[1];
const proto = el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, v);
el.dispatchEvent(new Event('input', {bubbles: true}));
el.dispatchEvent(new Event('change', {bubbles: true}));
"""

JS_VISIBLE_TEXT = JS_VIS + """
return Array.from(document.querySelectorAll(arguments[0])).filter(vis)
  .map(e => (e.innerText || '').trim()).filter(Boolean).join(' | ').slice(0, 300);
"""
JS_ANY_VISIBLE = JS_VIS + "return Array.from(document.querySelectorAll(arguments[0])).some(vis);"

# 网络请求监控：给 XHR / fetch 计数。点「查询」「导出」后看到请求发出又全部返回，就能确定数据已到。
# 忽略在线客服 / 心跳 / 埋点类请求：它们会一直挂着，旧版因此可能一直判定「等待服务器返回」直到超时。
JS_NET_HOOK = """
if (window.__yz) return true;
const yz = window.__yz = {pending: 0, started: 0, done: 0, last: [], urls: []};
const SKIP = /chat|kefu|customer-?service|websocket|socket\\.io|heartbeat|long-?poll|sentry|\\/track|beacon|hm\\.baidu|growingio|umeng|arms-retcode/i;
const FILE_RE = /\\.(xlsx?|csv|zip)(\\?|$)|export|download/i, IMG_RE = /\\.(png|jpe?g|gif|svg|webp|ico)(\\?|$)/i;
const grab = t => { try { if (typeof t !== 'string' || t.length > 200000) return;
  for (const v of (t.match(/https?:\\/\\/[^"'\\s<>\\\\]+/g) || []))
    if (FILE_RE.test(v) && !IMG_RE.test(v) && !yz.urls.includes(v)) yz.urls.push(v); } catch (e) {} };
const rec = (u, st, t0, ct, cd) => { yz.done++;
  yz.last.push({u: String(u || '').slice(0, 160), s: st, ms: Date.now() - t0, ct: (ct || '').slice(0, 60), cd: (cd || '').slice(0, 120)});
  if (yz.last.length > 30) yz.last.shift(); };
const XO = XMLHttpRequest.prototype.open, XS = XMLHttpRequest.prototype.send;
XMLHttpRequest.prototype.open = function (m, u) { this.__yzu = u; return XO.apply(this, arguments); };
XMLHttpRequest.prototype.send = function () {
  const t0 = Date.now(), x = this; let fin = false;
  if (SKIP.test(String(x.__yzu || ''))) return XS.apply(this, arguments);
  yz.pending++; yz.started++;
  x.addEventListener('loadend', () => { if (fin) return; fin = true;
    yz.pending = Math.max(0, yz.pending - 1);
    let ct = '', cd = '';
    try { ct = x.getResponseHeader('content-type') || ''; cd = x.getResponseHeader('content-disposition') || ''; } catch (e) {}
    try { if (x.responseType === '' || x.responseType === 'text') grab(x.responseText);
          else if (x.responseType === 'json') grab(JSON.stringify(x.response)); } catch (e) {}
    rec(x.__yzu, x.status, t0, ct, cd); });
  return XS.apply(this, arguments);
};
if (window.fetch) {
  const F = window.fetch;
  window.fetch = function (i) {
    const t0 = Date.now(), u = (i && i.url) || i;
    if (SKIP.test(String(u || ''))) return F.apply(window, arguments);
    yz.pending++; yz.started++;
    return F.apply(window, arguments).then(
      r => { yz.pending = Math.max(0, yz.pending - 1);
             const ct = r.headers.get('content-type') || '';
             rec(u, r.status, t0, ct, r.headers.get('content-disposition') || '');
             if (/json|text/i.test(ct)) { try { r.clone().text().then(grab); } catch (e) {} }
             return r; },
      e => { yz.pending = Math.max(0, yz.pending - 1); rec(u, 0, t0); throw e; });
  };
}
return true;
"""
JS_NET_RESET = ("const y = window.__yz; if (!y) return false; "
                "y.started = 0; y.done = 0; y.last = []; y.urls = []; return true;")
JS_NET_STATE = "const y = window.__yz; return y ? {p: y.pending, s: y.started, d: y.done} : null;"
JS_NET_DUMP = "const y = window.__yz; return y ? {last: y.last.slice(-8), urls: y.urls.slice(-5)} : null;"
JS_NET_PICK = """
const y = window.__yz; if (!y || !y.last.length) return null;
if (arguments[0] === 'last') return y.last[y.last.length - 1];
let m = y.last[0]; for (const x of y.last) if (x.ms > m.ms) m = x; return m;
"""


def net_hook(driver):
    return bool(_js(driver, JS_NET_HOOK, default=False))


def net_reset(driver):
    _js(driver, JS_NET_RESET)


def net_state(driver):
    return _js(driver, JS_NET_STATE)


def net_pick(driver, which="slowest"):
    return _js(driver, JS_NET_PICK, which)


def net_dump(driver):
    return _js(driver, JS_NET_DUMP) or {}


def spinning(driver, scope=None):
    return bool(_js(driver, JS_SPIN_IN, scope, SPIN_CSS, default=False))


def wait_loading_done(driver, timeout=None, scope=None):
    """等待加载遮罩消失。返回是否在超时前消失。"""
    return bool(wait_until(lambda: not spinning(driver, scope), timeout or CONFIG["TIMEOUT"], poll=0.3))


def click_text(driver, text, timeout=None):
    """点击文字为 text 的元素。优先按钮（兼容 antd 两字按钮自动加空格，如「查 询」），
    再精确文本，3 秒后仍找不到才用包含匹配（避免误点「母猪综合查询」这类标题）。"""
    timeout = timeout or CONFIG["TIMEOUT"]
    t0 = time.time()

    def find():
        for xp in (f"//button[translate(normalize-space(.),' ','')='{text}']",
                   f"//*[normalize-space(text())='{text}']"):
            for e in driver.find_elements(By.XPATH, xp):
                if e.is_displayed() and e.is_enabled():
                    return e
        if time.time() - t0 > 3:
            for e in driver.find_elements(By.XPATH, f"//*[contains(text(),'{text}')]"):
                t = e.text.strip()
                if t and len(t) <= len(text) + 10 and e.is_displayed():
                    return e
        return None

    el = wait_until(find, timeout, poll=0.3)
    if el is None:
        raise RuntimeError(f"未能在页面上找到可点击的「{text}」，请检查当前页面是否正确。")
    safe_click(driver, el)


# 顶部菜单项（antd Menu）：只认看得见的那一个（antd 会渲染隐藏的同名副本用来计算宽度）
JS_FIND_MENU = JS_VIS + """
const t = arguments[0];
for (const li of document.querySelectorAll("li[role='menuitem'], .ant-menu-item, .ant-menu-submenu-title")) {
  if (li.closest("[aria-hidden='true']")) continue;
  if (((li.getAttribute('title') || '') === t || norm(li.textContent) === t) && vis(li)) return li;
}
return null;
"""


def find_menu(driver, title):
    return _js(driver, JS_FIND_MENU, title)


# =====================================================================
# 窗口
# =====================================================================
def _handles(driver):
    try:
        return list(driver.window_handles)
    except WebDriverException as e:
        _raise_if_dead(e)
        return []


def _current_handle(driver):
    try:
        return driver.current_window_handle
    except WebDriverException as e:
        _raise_if_dead(e)
        return None       # 当前窗口已被关闭


def switch_to_new_window(driver, before, wait=1.5):
    """点击后若打开了新窗口（与点击前的窗口列表对比），切换过去。"""
    new = wait_until(lambda: [h for h in _handles(driver) if h not in before], wait, poll=0.2)
    if new:
        driver.switch_to.window(new[-1])
        log("已切换到新打开的页面")
        return True
    return False


def handle_new_window(driver):
    hs = _handles(driver)
    if len(hs) > 1 and _current_handle(driver) != hs[-1]:
        driver.switch_to.window(hs[-1])
        log("已切换到新打开的页面")


def close_other_windows(driver, keep=None):
    """只保留一个窗口（keep 不存在时保留第一个），返回保留的窗口句柄。"""
    hs = _handles(driver)
    if not hs:
        raise InvalidSessionIdException("浏览器窗口已全部关闭")
    keep = keep if keep in hs else hs[0]
    for h in hs:
        if h != keep:
            try:
                driver.switch_to.window(h)
                driver.close()
            except WebDriverException as e:
                _raise_if_dead(e)
    driver.switch_to.window(keep)
    return keep


def dismiss_guide(driver):
    """登录后若出现引导浮层，尽量点掉。"""
    _js(driver, JS_CLICK_BY_TEXT, ["跳过", "知道了", "我知道了", "暂不", "关闭"], None, "button, a, span")


CONFIRM_SCOPE = (".ant-modal-wrap, .ant-modal-confirm, .ant-popover, .el-message-box__wrapper, "
                 ".el-dialog__wrapper, [role='dialog']")


def confirm_dialog_if_any(driver, scope=CONFIRM_SCOPE):
    """导出后若弹出确认框，自动点确认（只在弹窗内找按钮，避免误点页面其它「确定」）。"""
    return bool(_js(driver, JS_CLICK_BY_TEXT, ["确认导出", "确定", "确认", "知道了"], scope, "button",
                    default=None))


# =====================================================================
# 页面右上角的「公司 / 场」选择框（母猪综合查询、母猪档案页面共用同一个组件）
# =====================================================================
JS_FARM_INPUT = JS_VIS + """
for (const sel of ["[class*='org-farm-selector-index-wrapper']", "[class*='org-farm-selector']", "[class*='farm-selector']"]) {
  for (const w of document.querySelectorAll(sel)) {
    if (!vis(w)) continue;
    const i = w.tagName === 'INPUT' ? w : w.querySelector('input');
    if (i && vis(i)) return i;
  }
}
return null;
"""

POPUP_CSS = ("[class*='org-farm-selector'][class*='popupWrap'], [class*='farm-selector'] [class*='popup'], "
             ".ant-popover, .ant-dropdown:not(.ant-dropdown-hidden), .ant-select-dropdown, "
             ".ant-cascader-dropdown, .ant-tree-select-dropdown, .ant-modal-wrap, .ant-drawer-open")

JS_POPUP_OPEN = JS_VIS + """
return Array.from(document.querySelectorAll(arguments[0])).some(r => vis(r) && norm(r.textContent).length > 0);
"""

JS_POPUP_SEARCH = JS_VIS + """
const main = arguments[1];
for (const r of Array.from(document.querySelectorAll(arguments[0])).filter(vis))
  for (const i of r.querySelectorAll("input[type='text'], input[type='search'], input:not([type])"))
    if (i !== main && vis(i) && !i.readOnly) return i;
return null;
"""

JS_FARM_OPTION = JS_VIS + """
const name = norm(arguments[1]), exact = arguments[2];
const hits = [];
for (const r of Array.from(document.querySelectorAll(arguments[0])).filter(vis)) {
  for (const e of r.querySelectorAll('*')) {
    if (e.tagName === 'INPUT' || !vis(e)) continue;
    const t = norm(e.textContent);
    if (!t) continue;
    const ok = exact ? t === name : (t.includes(name) && t.length <= name.length + 12);
    if (!ok) continue;
    if (Array.from(e.children).some(c => norm(c.textContent) === t)) continue;   // 取最内层
    if (!hits.includes(e)) hits.push(e);
  }
}
return hits;
"""

JS_POPUP_TEXTS = JS_VIS + """
const out = [];
for (const r of Array.from(document.querySelectorAll(arguments[0])).filter(vis))
  for (const e of r.querySelectorAll('li, [role=option], .ant-tree-title, .ant-select-item, span, div'))
    if (vis(e) && !e.children.length) { const t = norm(e.textContent); if (t && !out.includes(t)) out.push(t); }
return out.slice(0, 40);
"""



def escape_key(driver):
    """关掉可能残留的下拉 / 弹出层。"""
    try:
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    except Exception:
        pass


def current_org(driver):
    """右上角场选择框当前显示的公司 / 场名。"""
    inp = _js(driver, JS_FARM_INPUT)
    try:
        return (inp.get_attribute("value") or "") if inp is not None else ""
    except WebDriverException as e:
        _raise_if_dead(e)
        return ""


def _type_into(el, text):
    try:
        el.click()
    except Exception:
        pass
    el.send_keys(Keys.CONTROL, "a")
    el.send_keys(Keys.BACKSPACE)
    el.send_keys(text)


def select_org(driver, name):
    """右上角场选择框选中指定公司 / 场（需在该组件所在的顶层页面调用）。
    返回 True = 发生了切换；False = 本来就是它。找不到时抛出错误并列出可见的名称。"""
    target = _n(name)
    inp = wait_until(lambda: _js(driver, JS_FARM_INPUT), 15)
    if inp is None:
        save_debug(driver, "未找到场选择框")
        raise RuntimeError("页面右上角未找到公司 / 场选择框")
    cur = current_org(driver)
    if _n(cur) == target:
        return False

    log(f"切换公司 / 场：{cur or '（空）'} → {name}")
    net_hook(driver)
    net_reset(driver)
    safe_click(driver, inp)
    wait_until(lambda: _js(driver, JS_POPUP_OPEN, POPUP_CSS, default=False), 5, poll=0.2)

    search = _js(driver, JS_POPUP_SEARCH, POPUP_CSS, inp)     # 弹出层有搜索框就先搜索，列表长时更稳
    if search is not None:
        try:
            _type_into(search, name)
        except Exception:
            pass
    opts = wait_until(lambda: _js(driver, JS_FARM_OPTION, POPUP_CSS, name, True, default=[]), 6, poll=0.25)
    if not opts and search is None and not inp.get_attribute("readonly"):
        try:
            _type_into(inp, name)
        except Exception:
            pass
        opts = wait_until(lambda: _js(driver, JS_FARM_OPTION, POPUP_CSS, name, True, default=[]), 6, poll=0.25)
    if not opts:
        part = _js(driver, JS_FARM_OPTION, POPUP_CSS, name, False, default=[]) or []
        if len(part) == 1:
            opts = part
    if not opts:
        seen = _js(driver, JS_POPUP_TEXTS, POPUP_CSS, default=[]) or []
        save_debug(driver, f"未找到场_{name}")
        escape_key(driver)
        raise RuntimeError(f"公司 / 场选择列表里没有找到「{name}」。列表中可见的名称："
                           f"{'、'.join(seen[:20]) or '（未识别到弹出列表）'}")
    try:
        opts[0].click()
    except WebDriverException as e:
        _raise_if_dead(e)
        driver.execute_script("arguments[0].click();", opts[0])
    _js(driver, JS_CLICK_BY_TEXT, ["确定", "确认"], POPUP_CSS, "button")

    if not wait_until(lambda: _n(current_org(driver)) == target, 8, poll=0.25):
        save_debug(driver, f"切换场失败_{name}")
        escape_key(driver)
        raise RuntimeError(f"点选「{name}」后选择框显示为「{current_org(driver)}」，切换未生效")
    escape_key(driver)
    return True


def page_snapshot(driver):
    try:
        body = driver.find_element(By.TAG_NAME, "body").text[:300].replace("\n", " | ")
        log(f"当前 URL: {driver.current_url}")
        log(f"页面文本片段: {body}")
    except Exception:
        pass


def rotate_file(path, max_bytes=5 * 1024 * 1024, keep=3):
    """日志超过 max_bytes 时改名为 .1（原 .1 → .2 …），最多保留 keep 份旧日志。"""
    try:
        if not os.path.isfile(path) or os.path.getsize(path) < max_bytes:
            return False
        for i in range(keep, 0, -1):
            src = path if i == 1 else f"{path}.{i - 1}"
            dst = f"{path}.{i}"
            if os.path.exists(src):
                if os.path.exists(dst):
                    os.remove(dst)
                os.replace(src, dst)
        return True
    except OSError:
        return False


def clean_debug(days=30, max_files=200):
    """删掉 debug 文件夹里超过 days 天的出错截图 / 页面源码；文件太多时只保留最新的 max_files 个。"""
    d = CONFIG.get("DEBUG_DIR") or os.path.join(_app_dir(), "debug")
    try:
        files = [os.path.join(d, f) for f in os.listdir(d)]
    except OSError:
        return 0
    files = [f for f in files if os.path.isfile(f)]
    files.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    cutoff = time.time() - days * 86400
    n = 0
    for i, f in enumerate(files):
        try:
            if i >= max_files or os.path.getmtime(f) < cutoff:
                os.remove(f)
                n += 1
        except OSError:
            pass
    return n


def save_debug(driver, tag):
    """出错时保存截图和页面源码（后台运行时看不到浏览器，这是排查的主要依据）。"""
    try:
        d = CONFIG.get("DEBUG_DIR") or os.path.join(_app_dir(), "debug")
        os.makedirs(d, exist_ok=True)
        base = os.path.join(d, f"{time.strftime('%Y%m%d_%H%M%S')}_{sanitize(tag)[:40]}")
        driver.save_screenshot(base + ".png")
        with open(base + ".html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        log(f"已保存出错现场：{base}.png")
    except Exception:
        pass


# =====================================================================
# 账号 / 浏览器
# =====================================================================
def env_value(key):
    """读取设置项：先看环境变量，再看程序目录的 .env 文件。没有返回空字符串。"""
    v = os.environ.get(key, "").strip()
    if v:
        return v
    env_path = os.path.join(_app_dir(), ".env")
    try:
        with open(env_path, "r", encoding="utf-8-sig") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, val = line.split("=", 1)
                    if k.strip() == key:
                        return val.strip().strip('"').strip("'")
    except OSError:
        pass
    return ""


def load_credentials():
    """从环境变量读取账号密码；缺省时读取 exe/脚本同目录的 .env 文件。"""
    user = os.environ.get("YINGZI_USERNAME", "").strip()
    pwd = os.environ.get("YINGZI_PASSWORD", "").strip()
    env_path = os.path.join(_app_dir(), ".env")
    if (not user or not pwd) and os.path.exists(env_path):
        with open(env_path, "r", encoding="utf-8-sig") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    k, v = k.strip(), v.strip().strip('"').strip("'")
                    if k == "YINGZI_USERNAME" and not user:
                        user = v
                    elif k == "YINGZI_PASSWORD" and not pwd:
                        pwd = v
    if not user or not pwd:
        raise SystemExit(
            "未找到账号/密码。请设置环境变量 YINGZI_USERNAME / YINGZI_PASSWORD，\n"
            "或在程序同目录创建 .env 文件（YINGZI_USERNAME=账号 / YINGZI_PASSWORD=密码）。")
    return user, pwd


CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    r"C:\Program Files\Google\Chrome for Testing\Application\chrome.exe",
]
EDGE_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]
_UA_CACHE = {}


def _build_options(kind, binary, download_dir, user_agent=None, profile_dir=None):
    options = ChromeOptions() if kind == "chrome" else EdgeOptions()
    options.binary_location = binary
    options.page_load_strategy = "eager"     # DOM 就绪即继续，关键交互都有显式等待
    options.add_experimental_option("prefs", {
        "download.default_directory": os.path.abspath(download_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "safebrowsing.disable_download_protection": True,
        "download_restrictions": 0,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "profile.default_content_settings.popups": 0,
        "intl.accept_languages": "zh-CN,zh",
    })
    options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    args = ["--lang=zh-CN", "--no-first-run", "--no-default-browser-check",
            "--disable-notifications", "--disable-blink-features=AutomationControlled",
            # 不弹「不安全的下载」「下载气泡」等拦截提示，避免文件卡在浏览器里不落盘
            "--disable-features=InsecureDownloadWarnings,DownloadBubble,DownloadBubbleV2",
            "--safebrowsing-disable-download-protection"]
    if HEADLESS:
        # 无界面模式：强制中文、固定视口（否则日历标题是英文、布局变窄导致点选失败）
        args += ["--headless=new", "--window-size=1920,1080", "--force-device-scale-factor=1", "--disable-gpu"]
    else:
        args += ["--start-maximized"]
    if user_agent:
        args.append(f"--user-agent={user_agent}")
    if profile_dir:
        args += [f"--user-data-dir={profile_dir}", "--hide-crash-restore-bubble",
                 "--disable-session-crashed-bubble"]
    for a in args:
        options.add_argument(a)
    return options


def _service(kind, path=None):
    Service = ChromeService if kind == "chrome" else EdgeService
    svc = Service(executable_path=path) if path else Service()
    if os.name == "nt":
        try:
            svc.creation_flags = 0x08000000   # CREATE_NO_WINDOW：不弹出 chromedriver 黑窗口
        except Exception:
            pass
    return svc


def _profile_root():
    return CONFIG.get("PROFILE_DIR") or os.path.join(_app_dir(), "browser_profile")


def _launch_once(kind, binary, download_dir, user_agent, profile):
    options = _build_options(kind, binary, download_dir, user_agent, profile)
    cls = ChromeDriver if kind == "chrome" else EdgeDriver
    try:
        return cls(service=_service(kind), options=options)
    except Exception as e:
        if profile and "user data" in str(e).lower():
            raise
        log(f"启动浏览器失败（{str(e)[:80]}），改用 webdriver-manager ...")
        if kind == "chrome":
            from webdriver_manager.chrome import ChromeDriverManager
            path = ChromeDriverManager().install()
        else:
            from webdriver_manager.microsoft import EdgeChromiumDriverManager
            path = EdgeChromiumDriverManager().install()
        return cls(service=_service(kind, path), options=options)


def _launch(kind, binary, download_dir, user_agent=None):
    profile = None
    if CONFIG.get("KEEP_LOGIN"):
        profile = os.path.join(_profile_root(), kind)
        try:
            os.makedirs(profile, exist_ok=True)
        except OSError:
            profile = None
    try:
        return _launch_once(kind, binary, download_dir, user_agent, profile)
    except Exception as e:
        if not profile:
            raise
        log(f"带登录缓存启动失败（{str(e)[:60]}；可能另一个下载器正在运行），改为不保留登录状态 ...")
        return _launch_once(kind, binary, download_dir, user_agent, None)


def allow_downloads(driver, download_dir):
    """显式允许下载到指定目录（无界面 Chrome 不一定遵循 prefs，必须用 CDP 设置）。"""
    params = {"behavior": "allow", "downloadPath": os.path.abspath(download_dir)}
    for cmd in ("Browser.setDownloadBehavior", "Page.setDownloadBehavior"):
        try:
            driver.execute_cdp_cmd(cmd, params)
            SESSION["current_dir"] = params["downloadPath"]   # 浏览器当前实际的下载目录
            return True
        except Exception:
            continue
    log("提示：未能通过 CDP 设置下载目录，将依赖浏览器默认设置")
    return False


def create_driver(download_dir):
    """创建浏览器（自动检测 Chrome / Edge）。"""
    global CURRENT_DRIVER
    force_quit()      # 上一个浏览器没关就开新的，会锁住登录缓存目录、留下后台进程
    os.makedirs(download_dir, exist_ok=True)
    chrome_bin = _first_existing(
        [CONFIG.get("CHROME_PATH", ""), os.environ.get("YINGZI_CHROME_PATH", "")] + CHROME_CANDIDATES)
    edge_bin = _first_existing([os.environ.get("YINGZI_EDGE_PATH", "")] + EDGE_CANDIDATES)
    if chrome_bin:
        kind, binary = "chrome", chrome_bin
    elif edge_bin:
        kind, binary = "edge", edge_bin
    else:
        raise SystemExit("未找到 Chrome 或 Edge 浏览器。请安装 Chrome，\n"
                         "或设置环境变量 YINGZI_CHROME_PATH 指向 chrome.exe。")

    log(f"启动浏览器（{'后台运行' if HEADLESS else '显示窗口'}）...")
    driver = _launch(kind, binary, download_dir, _UA_CACHE.get(kind) if HEADLESS else None)
    if HEADLESS:
        try:
            ua = driver.execute_script("return navigator.userAgent") or ""
        except Exception:
            ua = ""
        if "Headless" in ua:
            # 部分网站会拦截 HeadlessChrome，用普通 UA 重启一次（仅首次）
            _UA_CACHE[kind] = ua.replace("HeadlessChrome", "Chrome")
            try:
                driver.quit()
            except Exception:
                pass
            driver = _launch(kind, binary, download_dir, _UA_CACHE[kind])
    CURRENT_DRIVER = driver
    SESSION["browser_dir"] = os.path.abspath(download_dir)   # 浏览器启动时的默认下载目录
    allow_downloads(driver, download_dir)
    driver.set_page_load_timeout(60)
    return driver


def force_quit():
    """关闭浏览器，避免后台残留 Chrome 进程。"""
    global CURRENT_DRIVER
    d, CURRENT_DRIVER = CURRENT_DRIVER, None
    if d is not None:
        try:
            d.quit()
        except Exception:
            pass


# =====================================================================
# 登录
# =====================================================================
LOGIN_MSG_CSS = (".ant-message-notice, .ant-notification-notice, .ant-form-item-explain-error, .ant-form-explain, "
                 ".ant-alert-error, .el-message, .el-form-item__error, [class*='errorMsg'], [class*='error-msg'], "
                 "[class*='error-tip']")
CAPTCHA_CSS = ("[class*='captcha'], [id*='captcha'], [class*='geetest'], [class*='nc_wrapper'], "
               "[class*='slide-verify'], [class*='verifybox'], iframe[src*='captcha']")
HOME_XPATH = ("//*[contains(text(),'养殖生产') or contains(text(),'养殖公共') or contains(text(),'统计报表') "
              "or contains(text(),'个体养殖') or contains(text(),'报表名称')]")


def _home_url_file():
    return os.path.join(_profile_root(), "home_url.txt")


def _load_home_url():
    try:
        with open(_home_url_file(), encoding="utf-8") as f:
            return f.read().strip()
    except OSError:
        return ""


def _save_home_url(url):
    try:
        os.makedirs(_profile_root(), exist_ok=True)
        with open(_home_url_file(), "w", encoding="utf-8") as f:
            f.write(url)
    except OSError:
        pass


def open_url(driver, url, tries=3):
    """打开网址。页面加载超时不算失败（单页应用常有慢资源），停止加载后继续操作。"""
    for attempt in range(tries):
        check_stop()
        try:
            driver.get(url)
            return
        except TimeoutException:
            _js(driver, "window.stop();")
            log("页面加载较慢，已停止等待剩余资源，继续操作")
            return
        except WebDriverException as e:
            _raise_if_dead(e)
            if attempt == tries - 1:
                raise
            wait = 3 * (attempt + 1)
            log(f"打开页面失败（{str(e)[:60]}），{wait} 秒后重试 ...")
            time.sleep(wait)


def check_agreement(driver):
    """勾选「已阅读并同意」。只点 checkbox 本身，避免误触协议链接。"""
    xpath = "//input[@type='checkbox' and ancestor::label[contains(.,'已阅读')]]"
    if not wait_until(lambda: driver.find_elements(By.XPATH, xpath), 5):
        return
    for _ in range(5):
        try:
            cb = driver.find_element(By.XPATH, xpath)
            if not cb.is_selected():
                driver.execute_script("arguments[0].click();", cb)
            if wait_until(lambda: driver.find_element(By.XPATH, xpath).is_selected(), 1.5, poll=0.15):
                return
        except (StaleElementReferenceException, NoSuchElementException):
            time.sleep(0.3)
    log("勾选协议未成功（继续尝试登录）")


def _fill_input(driver, el_id, value):
    """填写输入框并核对；普通输入没生效时，用 React 兼容的方式直接写值。"""
    for _ in range(3):
        try:
            el = driver.find_element(By.ID, el_id)
            el.click()
            el.send_keys(Keys.CONTROL, "a")
            el.send_keys(Keys.BACKSPACE)
            el.send_keys(value)
            if (el.get_attribute("value") or "") == value:
                return True
            driver.execute_script(JS_SET_VALUE, el, value)
            if (el.get_attribute("value") or "") == value:
                return True
        except (StaleElementReferenceException, NoSuchElementException,
                ElementNotInteractableException, ElementClickInterceptedException):
            pass
        time.sleep(0.3)
    return False


def _at_home(driver):
    """已进入系统：离开登录页，且能看到「养殖生产」「养殖公共」「统计报表」等菜单。"""
    try:
        driver.switch_to.default_content()
    except WebDriverException as e:
        _raise_if_dead(e)
    if "user/login" in driver.current_url:
        return False
    return bool(driver.find_elements(By.XPATH, HOME_XPATH))


def _login_form_visible(driver):
    return any(e.is_displayed() for e in driver.find_elements(By.ID, "username"))


def _login_once(driver, username, password):
    wait_loading_done(driver, 10)
    if not (_fill_input(driver, "username", username) and _fill_input(driver, "password", password)):
        raise RuntimeError("账号/密码未能填入输入框")
    check_agreement(driver)
    # 勾选协议可能触发页面重渲染、清空输入，再核对一次
    for el_id, val in (("username", username), ("password", password)):
        try:
            cur = driver.find_element(By.ID, el_id).get_attribute("value") or ""
        except NoSuchElementException:
            cur = ""
        if cur != val:
            _fill_input(driver, el_id, val)
    log("已填写账号密码")

    def find_btn():
        for b in driver.find_elements(By.XPATH, "//button[@type='submit' or contains(.,'登')]"):
            if b.is_displayed() and b.is_enabled() and "登录" in b.text.replace(" ", ""):
                return b
        return None

    btn = wait_until(find_btn, 10)
    if btn is None:
        check_agreement(driver)
        btn = wait_until(find_btn, 5)
    if btn is None:
        raise RuntimeError("登录按钮一直不可用（可能协议未勾选）")
    safe_click(driver, btn)
    log("已点击登录，等待进入系统 ...")

    t0 = time.time()
    deadline = t0 + 45
    pressed_enter = captcha = captcha_logged = False
    shown = set()
    while True:
        check_stop()
        now = time.time()
        if "user/login" not in driver.current_url:
            return
        msg = _js(driver, JS_VISIBLE_TEXT, LOGIN_MSG_CSS, default="") or ""
        if msg:
            if re.search(r"密码错误|账号或密码|用户名或密码|密码不正确|不存在|锁定|冻结|禁用|停用", msg):
                raise LoginFatal(f"登录失败：{msg}")
            if re.search(r"请输入|请勾选|请阅读", msg):
                raise RuntimeError(f"登录表单未填写完整：{msg}")
            if re.search(r"验证", msg):
                captcha = True
            elif re.search(r"异常|失败|错误|超时|繁忙", msg):
                raise RuntimeError(f"登录提示：{msg}")
            elif msg not in shown:
                shown.add(msg)
                log(f"登录页提示：{msg}")
        if not captcha and _js(driver, JS_ANY_VISIBLE, CAPTCHA_CSS, default=False):
            captcha = True
        if captcha:
            if HEADLESS:
                raise LoginFatal("登录需要验证码。请取消勾选「后台运行」，在弹出的浏览器里手动完成一次验证；"
                                 "勾选「记住登录状态」后，之后即可免登录、后台运行。")
            if not captcha_logged:
                captcha_logged = True
                deadline = now + CONFIG["CAPTCHA_WAIT"]
                log(f"检测到验证码，请在浏览器窗口中手动完成验证（最多等待 {CONFIG['CAPTCHA_WAIT']} 秒）...")
        if not pressed_enter and not captcha and now - t0 > 8:
            # 点击后迟迟没反应（按钮事件未绑定好等），再用回车提交一次
            try:
                driver.find_element(By.ID, "password").send_keys(Keys.ENTER)
            except Exception:
                pass
            pressed_enter = True
        if now > deadline:
            raise RuntimeError("点击登录后长时间未进入系统")
        time.sleep(0.4)


def login(driver, username, password):
    """登录（带自动重试）。记住登录状态时先尝试直接打开上次的首页，有效就跳过登录。"""
    if CONFIG.get("KEEP_LOGIN"):
        home = _load_home_url()
        if home:
            log("尝试使用已保存的登录状态 ...")
            open_url(driver, home)
            if wait_until(lambda: _at_home(driver), 12):
                log("登录状态有效，跳过登录")
                SESSION["home_url"] = driver.current_url
                dismiss_guide(driver)
                return
            log("登录状态已失效，重新登录")

    last = None
    for attempt in range(1, CONFIG["LOGIN_ATTEMPTS"] + 1):
        try:
            log("打开登录页 ..." if attempt == 1 else f"第 {attempt} 次尝试登录 ...")
            open_url(driver, CONFIG["BASE_URL"])
            state = wait_until(lambda: "form" if _login_form_visible(driver)
                               else ("in" if _at_home(driver) else None), CONFIG["TIMEOUT"])
            if state is None:
                raise RuntimeError("登录页加载超时")
            if state == "form":
                _login_once(driver, username, password)
            last = None
            break
        except LoginFatal:
            save_debug(driver, "登录失败")
            raise
        except InvalidSessionIdException:
            raise
        except Exception as e:
            _raise_if_dead(e)
            last = e
            log(f"登录未成功：{str(e)[:120]}")
            save_debug(driver, f"登录失败_第{attempt}次")
            time.sleep(2)
    if last is not None:
        raise RuntimeError(f"连续 {CONFIG['LOGIN_ATTEMPTS']} 次登录失败：{last}")

    handle_new_window(driver)
    wait_until(lambda: _at_home(driver), 20)     # 等首页菜单出现，后续导航更稳
    dismiss_guide(driver)
    if _at_home(driver):
        SESSION["home_url"] = driver.current_url
        if CONFIG.get("KEEP_LOGIN"):
            _save_home_url(driver.current_url)
    log("登录成功")


def start_session(download_dir, username, password, goto):
    """启动浏览器 → 登录 → goto(driver) 进入目标页面。中途出错会关掉浏览器再抛出（旧版会残留后台 Chrome）。"""
    SESSION["home_url"] = None
    driver = create_driver(download_dir)
    try:
        login(driver, username, password)
        goto(driver)
        return driver
    except BaseException:
        force_quit()
        raise


def back_to_home(driver, keep=None):
    """不重新登录，只保留一个窗口并回到首页（清掉页面上残留的弹窗、状态）。"""
    keep = close_other_windows(driver, keep)
    if SESSION.get("home_url"):
        open_url(driver, SESSION["home_url"])
    else:
        try:
            driver.refresh()
        except TimeoutException:
            _js(driver, "window.stop();")
    if "user/login" in driver.current_url:
        raise RuntimeError("登录已失效")
    wait_until(lambda: _at_home(driver), 30)
    return keep


# =====================================================================
# 查询 / 表格
# =====================================================================
JS_TABLE_STATE = JS_VIS + """
const q = s => Array.from(document.querySelectorAll(s));
const spin = q(arguments[0]).some(vis);
const rows = q('table tbody tr').filter(vis);
const head = rows.slice(0, 3).map(r => r.textContent).join('|').slice(0, 600);
const cells = q('table tbody td').filter(e => vis(e) && e.textContent.trim()).length;
const empty = q('.ant-empty, .ant-table-placeholder, .el-table__empty-block, .davinci-table-empty, .davinci-empty').some(vis);
const total = q('.ant-pagination-total-text, .el-pagination__total, .davinci-pagination-total-text').filter(vis).map(e => e.textContent).join('|');
let exp = null;
for (const b of q('button')) {
  if (norm(b.textContent).includes('导出') && vis(b)) {
    exp = !(b.disabled || b.classList.contains('is-disabled')); break; }
}
const yz = window.__yz;
return {spin: spin, sig: rows.length + '#' + head + '#' + total, empty: empty, cells: cells, exp: exp,
        net: yz ? {p: yz.pending, s: yz.started, d: yz.done} : null,
        msg: q('.ant-message-error, .ant-message-warning, .ant-notification-notice-error, .ant-notification-notice-warning, .el-message--error, .el-message--warning')
               .filter(vis).map(e => (e.innerText || '').trim()).filter(Boolean).join(' | ').slice(0, 200)};
"""

# 分页里的「共 N 条」（不看表格内容和弹窗 / 抽屉里的分页）
JS_TOTAL_ROWS = JS_VIS + """
const bad = e => e.closest('tbody, .ant-drawer, .ant-modal-wrap, .el-dialog__wrapper, .el-drawer__wrapper');
const num = e => { const t = (e.textContent || '');
  if (t.length > 40) return null;
  const m = t.replace(/\\s+/g, '').match(/共([\\d,]+)条/);
  if (!m) return null; const n = parseInt(m[1].replace(/,/g, ''), 10); return isNaN(n) ? null : n; };
// 优先读分页组件的「共 N 条」；页面上有多个时取第一个看得见的（旧版取最大值，可能读到别处的数字）
for (const e of document.querySelectorAll('.davinci-pagination-total-text, .ant-pagination-total-text, .el-pagination__total')) {
  if (vis(e) && !bad(e)) { const n = num(e); if (n !== null) return n; }
}
// 兜底：只看「叶子」元素（自身文字就是「共 N 条」，不是包着整块区域的大容器）
for (const e of document.querySelectorAll('li, span, div, p')) {
  if (bad(e) || !vis(e) || Array.from(e.children).some(c => /共/.test(c.textContent || ''))) continue;
  const n = num(e); if (n !== null) return n;
}
return null;
"""

# 导出被拒绝的提示：只看错误/警告类消息（普通弹窗正文里的「单次最多导出5万条」说明不算）
EXPORT_ERR_CSS = (".ant-message-error, .ant-message-warning, .ant-notification-notice-error, "
                  ".ant-notification-notice-warning, .ant-modal-confirm-error, .ant-modal-confirm-warning, "
                  ".el-message--error, .el-message--warning, .el-notification .el-icon-error, "
                  ".el-notification .el-icon-warning")
LIMIT_NUM_RE = r"5\s*万|50000|50,000|五万|上限"
LIMIT_EXCEED_RE = r"超过|超出|大于|不能|不允许|无法|过多|太多|过大"


def is_limit_msg(m):
    return bool(m) and bool(re.search(LIMIT_NUM_RE, m)) and bool(re.search(LIMIT_EXCEED_RE, m))


def table_state(driver):
    return _js(driver, JS_TABLE_STATE, SPIN_CSS, default={}) or {}


def query_total(driver):
    v = _js(driver, JS_TOTAL_ROWS)
    try:
        return int(v) if v is not None else None
    except (TypeError, ValueError):
        return None


def wait_query_done(driver, before, t0, timeout, grace, state_fn=None,
                    err_re=r"失败|错误|异常|请选择|不能|只能|超出|超过", what="查询"):
    """判断查询何时完成：
       1) 优先看网络请求：请求发出后全部返回、且无加载动画，稳定片刻即完成；
       2) 拿不到网络信息时看页面：加载动画出现又消失 / 表格内容变化 / 出现空数据提示；
       3) 始终没有任何信号，grace 秒后按完成处理。"""
    state_fn = state_fn or table_state
    end = t0 + timeout
    seen_spin = net_seen = False
    idle_since, last_log, seen_msgs = None, t0, set()
    while time.time() < end:
        check_stop()
        s = state_fn(driver) or {}
        now = time.time()
        net = s.get("net")
        if net and net.get("s", 0) > 0:
            net_seen = True
        m = s.get("msg")
        if m and m not in seen_msgs:
            seen_msgs.add(m)
            if err_re and re.search(err_re, m):
                raise RuntimeError(f"{what}时页面提示：{m}")
            log(f"页面提示：{m}")
        busy = None
        if net and net.get("p", 0) > 0:
            busy = "等待服务器返回数据"
        elif s.get("spin"):
            busy = "页面加载中"
        elif s.get("exp") is False:
            busy = "导出按钮尚未可用"
        if busy:
            seen_spin = seen_spin or bool(s.get("spin"))
            idle_since = None
        else:
            changed = (s.get("sig") != before.get("sig")) or (s.get("empty") and not before.get("empty"))
            if net_seen or seen_spin or changed:
                idle_since = idle_since or now
                if now - idle_since >= (0.6 if net_seen else 1.0) and now - t0 >= (0.8 if net_seen else 1.5):
                    msg = f"{what}完成，用时 {fmt_secs(now - t0)}"
                    slow = net_pick(driver, "slowest") if net_seen else None
                    if slow and slow.get("ms", 0) >= 1000:
                        msg += f"（其中服务器处理 {slow['ms'] / 1000:.1f} 秒）"
                    if "cells" in s and not s.get("cells"):
                        msg += "，结果可能为空"
                    log(msg)
                    return True
            elif now - t0 >= grace:
                log(f"{grace} 秒内未检测到请求或数据变化，按{what}完成处理")
                return True
        if now - last_log >= 10:
            log(f"{what}中（{busy or '等待页面变化'}）... 已等待 {fmt_secs(now - t0)}")
            last_log = now
        time.sleep(0.3)
    log(f"等待{what}结果超时，继续后续步骤 ...")
    return False


# =====================================================================
# 下载目录
# =====================================================================
PARTIAL_EXT = (".crdownload", ".tmp", ".partial", ".download")


def watch_dirs(download_dir):
    """要监控的目录：下载目录 + 系统默认下载目录（部分浏览器配置下文件会落到那里）。"""
    out = []
    for d in [download_dir, SESSION.get("current_dir"), SESSION.get("browser_dir")] + list(CONFIG.get("EXTRA_DOWNLOAD_DIRS") or []):
        if d:
            d = os.path.abspath(d)
            if d not in out:
                out.append(d)
    return out


def list_files(dirs, exts=None):
    """目录（不含子目录）里已下载完成的文件。exts 为 None 时不限扩展名。"""
    out = set()
    for d in dirs:
        try:
            if not os.path.isdir(d):
                continue
            for f in os.listdir(d):
                low = f.lower()
                if low.endswith(PARTIAL_EXT) or f.startswith((".", "~$")) or low == "desktop.ini":
                    continue
                if exts and not low.endswith(tuple(exts)):
                    continue
                p = os.path.join(d, f)
                if os.path.isfile(p):
                    out.add(p)
        except OSError:
            pass
    return out


def recent_partials(dirs, window=30):
    """最近 window 秒内有变化的未完成下载文件（说明正在下载中）。"""
    now, out = time.time(), []
    for d in dirs:
        try:
            for f in os.listdir(d):
                if f.lower().endswith(PARTIAL_EXT):
                    p = os.path.join(d, f)
                    if now - os.path.getmtime(p) < window:
                        out.append(p)
        except OSError:
            pass
    return out


def file_stable(path, wait=0.5):
    try:
        a = os.path.getsize(path)
        time.sleep(wait)
        return a > 0 and os.path.getsize(path) == a
    except OSError:
        return False


def move_unique(src, folder, stem, ext=None):
    """把下载到的文件移动到 folder 并命名为 stem（重名自动加 _1 _2 …）。文件刚落盘时可能被占用，短暂重试。"""
    os.makedirs(folder, exist_ok=True)
    ext = ext or os.path.splitext(src)[1].lower() or ".xlsx"
    target = os.path.join(folder, stem + ext)
    n = 1
    while os.path.exists(target):
        target = os.path.join(folder, f"{stem}_{n}{ext}")
        n += 1
    for _ in range(20):
        try:
            shutil.move(src, target)
            return target
        except PermissionError:
            time.sleep(0.5)
    shutil.move(src, target)
    return target


def move_replace(src, folder, stem, ext=None):
    """把下载到的文件移动到 folder/stem.ext；同名文件已存在时直接覆盖（不产生 _1 _2）。"""
    os.makedirs(folder, exist_ok=True)
    ext = ext or os.path.splitext(src)[1].lower() or ".xlsx"
    target = os.path.join(folder, stem + ext)
    for i in range(20):
        try:
            if os.path.exists(target):
                os.remove(target)
            shutil.move(src, target)
            return target
        except PermissionError:
            if i == 0 and os.path.exists(target):
                log(f"提示：{os.path.basename(target)} 正被占用（可能用 Excel 打开着），稍后重试覆盖")
            time.sleep(0.5)
    raise RuntimeError(f"无法覆盖 {target}：文件被占用，请关闭 Excel 后重试")


# =====================================================================
# 按年 / 月分文件夹
# =====================================================================
def _to_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    nums = [int(x) for x in re.findall(r"\d+", str(v or ""))]
    return datetime.date(nums[0], nums[1], nums[2]) if len(nums) >= 3 else None


def period_folder(base, start=None, end=None):
    """按日期范围决定保存位置：
       同一个月（含整月、月内几天）→ base/2026年/06月
       同一年跨月（季度、全年……）  → base/2026年
       跨年                        → base/2025-2026年
       关闭 DATE_FOLDERS 或没有日期时返回 base。"""
    if not CONFIG.get("DATE_FOLDERS", True):
        return base
    try:
        a, b = _to_date(start), _to_date(end if end is not None else start)
    except (ValueError, TypeError):
        return base
    if a is None or b is None:
        return base
    if a > b:
        a, b = b, a
    if a.year != b.year:
        return os.path.join(base, f"{a.year}-{b.year}年")
    if a.month != b.month:
        return os.path.join(base, f"{a.year}年")
    return os.path.join(base, f"{a.year}年", f"{a.month:02d}月")


def rel_to(path, base):
    try:
        return os.path.relpath(path, base)
    except ValueError:
        return path


# =====================================================================
# 统计导出文件的数据行数（文件条数是最终依据）
# =====================================================================
_MERGE_RE = re.compile(rb'<(?:\w+:)?mergeCell\s+ref="[A-Z]+(\d+):[A-Z]+(\d+)"')


def _first_sheet_xml(zf):
    names = zf.namelist()
    try:   # workbook.xml 里第一个 sheet → rels → 文件路径
        wb = zf.read("xl/workbook.xml")
        rid = re.search(rb'<(?:\w+:)?sheet\b[^>]*\br:id="([^"]+)"', wb).group(1)
        rels = zf.read("xl/_rels/workbook.xml.rels")
        for m in re.finditer(rb"<Relationship\b[^>]*>", rels):
            tag = m.group(0)
            if b'Id="' + rid + b'"' in tag:
                t = re.search(rb'Target="([^"]+)"', tag).group(1).decode()
                t = t.lstrip("/")
                t = t if t.startswith("xl/") else "xl/" + t
                if t in names:
                    return t
    except Exception:
        pass
    cands = sorted(n for n in names if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n))
    return cands[0] if cands else None


def _xlsx_merges(path):
    """第一个工作表里所有合并单元格的 (起始行, 结束行)。只扫描工作表 XML 的末尾（mergeCells 在数据后面），大文件也很快。"""
    try:
        with zipfile.ZipFile(path) as zf:
            name = _first_sheet_xml(zf)
            if not name:
                return []
            tail = b""
            with zf.open(name) as f:
                while True:
                    chunk = f.read(1 << 20)
                    if not chunk:
                        break
                    tail = (tail + chunk)[-(4 << 20):]
            return [(int(a), int(b)) for a, b in _MERGE_RE.findall(tail)]
    except Exception:
        return []


def xlsx_data_rows(path):
    """统计 xlsx 第一个工作表的数据行数（去掉标题行、表头、空行）。读不了返回 None。
    表头判断：第 1 行只有一个单元格有字、下一行有多个 → 第 1 行是标题；
    表头从标题之后开始，被合并单元格（多级表头）向下延伸到哪一行，表头就到哪一行。"""
    if not str(path).lower().endswith((".xlsx", ".xlsm")):
        return None
    try:
        import openpyxl
    except ImportError:
        return None
    try:
        merges = _xlsx_merges(path)
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            filled = lambda r: sum(1 for v in r if v not in (None, ""))
            ws = wb.worksheets[0]
            # 平台（Java 导出库）写的表格范围常常是错的（只写 A1），快速读取模式会照着它只读第一行 → 数出 0 条。
            # reset_dimensions：忽略文件自报的范围，按实际内容读到最后一行。
            ws.reset_dimensions()
            rows = ws.iter_rows(values_only=True)
            head = []
            for r in rows:
                head.append(r)
                if len(head) >= 3:
                    break
            start = 1
            if len(head) >= 2 and filled(head[0]) == 1 and filled(head[1]) >= 3:
                start = 2                                   # 第 1 行是报表标题
            header_end = max([b for a, b in merges if a == start] + [start])
            n = sum(1 for i, r in enumerate(head, 1) if i > header_end and filled(r))
            for i, r in enumerate(rows, len(head) + 1):
                if i > header_end and filled(r):
                    n += 1
        finally:
            wb.close()
        if n == 0 and os.path.getsize(path) > 30 * 1024:
            # 文件不小却数出 0 条：换完整读取方式再数一遍，避免再出现「140 KB、0 条」
            wb = openpyxl.load_workbook(path, data_only=True)
            try:
                n = sum(1 for i, r in enumerate(wb.worksheets[0].iter_rows(values_only=True), 1)
                        if i > header_end and filled(r))
            finally:
                wb.close()
        return n
    except Exception:
        return None


def append_csv(path, header, row):
    """往记录表追加一行（不存在则带表头新建）；被 Excel 占用时重试几次，仍失败只提示不报错。"""
    for i in range(5):
        try:
            new = not os.path.exists(path)
            os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
            with open(path, "a", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                if new:
                    w.writerow(header)
                w.writerow(row)
            return True
        except PermissionError:
            time.sleep(1)
        except OSError:
            break
    log(f"警告：写入 {os.path.basename(path)} 失败（若正被 Excel 打开，请关闭后再试）")
    return False


# =====================================================================
# 通用任务循环
# =====================================================================
class Adapter:
    """一个下载功能的接口：
       goto(driver)                 已登录状态下，从首页（或任意页面）进入该功能页面
       run(driver, task) -> 结果    返回 None 表示「无数据」
       reload(driver)               不重新登录，重新打开该功能页面
       name(task) -> str
       download_dir() -> str        该功能的下载目录
       title                        功能名称（日志用）
       open_session(u, p)           （可选）启动浏览器 + 登录 + 进入页面；默认 = start_session(...) + goto
       begin()                      （可选）每次运行开始时调用"""

    def __init__(self, goto, run, reload, name, download_dir, title="", open_session=None, begin=None):
        self.goto, self.run, self.reload, self.name = goto, run, reload, name
        self.download_dir, self.title, self.begin = download_dir, title, begin
        self.open_session = open_session or (
            lambda u, p: start_session(self.download_dir(), u, p, self.goto))

    def enter(self, driver):
        """浏览器已登录（上一个功能刚用完）：回首页 → 切换下载目录 → 进入本功能页面，不重启浏览器。"""
        log(f"不重启浏览器，返回首页后进入「{self.title}」...")
        back_to_home(driver)
        if not allow_downloads(driver, self.download_dir()):
            log("提示：下载目录未能切换，文件可能先落到上一个功能的目录，程序会自动找到并移走")
        self.goto(driver)


def run_tasks(tasks, adapter, username=None, password=None, status_fn=None, what="任务", session=None):
    """逐个执行任务：失败先「重开页面」重试一次；浏览器崩溃 / 重开失败则重启浏览器重试。
    status_fn(i, total, state, info)，state: start / ok / empty / fail / stopped。返回成功结果列表。
    session：多个功能共用一个浏览器时传入同一个 dict（{"driver": ...}）。此时本函数结束后不关闭浏览器，
             下一个功能直接回首页进入，不用重新启动浏览器、重新登录；由调用方最后调用 force_quit()。
    账号密码错误等（LoginFatal）会在标记当前任务失败后抛出。"""
    def report(i, state, info=None):
        if status_fn:
            try:
                status_fn(i, len(tasks), state, info)
            except Exception:
                pass

    if not username or not password:
        username, password = load_credentials()
    if adapter.begin:
        adapter.begin()
    log(f"待下载{what}：{len(tasks)} 个")
    t_all = time.time()
    results = []
    own = session is None
    st = {"driver": None} if own else session
    if st.get("driver") is not None and CURRENT_DRIVER is not st["driver"]:
        st["driver"] = None                   # 共享的浏览器已被关掉（如强制停止）
    entered = {"ok": False}                   # 本功能页面是否已进入

    def dead():
        force_quit()
        st["driver"] = None

    def ensure():
        check_stop()
        if st["driver"] is not None and not entered["ok"]:
            try:
                adapter.enter(st["driver"])
                entered["ok"] = True
            except StopRequested:
                raise
            except Exception as e:
                log(f"回首页切换功能失败（{str(e)[:80]}），重启浏览器 ...")
                dead()
        if st["driver"] is None:
            st["driver"] = adapter.open_session(username, password)
            entered["ok"] = True
        return st["driver"]

    need_reload = False
    try:
        for i, task in enumerate(tasks, 1):
            if _stopped():
                log("已收到停止指令，停止后续任务")
                break
            name = adapter.name(task)
            log(f"进度：第 {i}/{len(tasks)} 个（{name}）")
            report(i, "start")
            err = None
            for rnd in range(2):
                try:
                    check_stop()
                    if (rnd == 1 or need_reload) and st["driver"] is not None and entered["ok"]:
                        need_reload = False
                        try:
                            adapter.reload(st["driver"])
                        except StopRequested:
                            raise
                        except Exception as e:
                            log(f"重新打开页面失败（{str(e)[:80]}），重启浏览器 ...")
                            dead()
                    res = adapter.run(ensure(), task)
                    err = None
                    if res is None:
                        report(i, "empty", "无数据")
                    else:
                        results.append(res)
                        report(i, "ok", res)
                    break
                except StopRequested:
                    report(i, "stopped")
                    raise
                except LoginFatal as e:
                    report(i, "fail", str(e)[:120])
                    raise
                except SystemExit as e:          # 未找到浏览器 / 账号等，重试也没用
                    report(i, "fail", str(e)[:120])
                    raise
                except Exception as e:
                    if _stopped():               # 「停止」时强制关闭浏览器引起的报错，不要再重开浏览器
                        report(i, "stopped")
                        raise StopRequested()
                    err = e
                    if is_dead_error(e):
                        log("浏览器会话中断（浏览器被关闭或崩溃）"
                            + ("，自动重启浏览器重试本任务 ..." if rnd == 0 else ""))
                        dead()
                    else:
                        log(f"「{name}」失败：{type(e).__name__}: {str(e)[:200]}")
                        if rnd == 0:
                            log("不退出登录，重新打开页面后重试本任务 ...")
            if err is not None:
                log(f"重试仍失败，跳过「{name}」")
                report(i, "fail", str(err)[:120])
                need_reload = True
    except StopRequested:
        log("用户已停止下载")
        if not own:
            raise
    finally:
        log("=" * 40)
        log(f"已完成 {len(results)}/{len(tasks)} 个{what}，总用时 {fmt_secs(time.time() - t_all)}")
        if own:
            force_quit()
    return results


def run_all(segments, username=None, password=None, status_fn=None):
    """多个功能连续执行、共用一个浏览器：功能之间只回首页，不重启浏览器、不重新登录。
    segments：[(adapter, tasks), ...]；status_fn(全局序号, 总数, state, info)。返回所有成功结果。"""
    if not username or not password:
        username, password = load_credentials()
    total = sum(len(t) for _, t in segments)
    session = {"driver": None}
    results, offset = [], 0
    try:
        for adapter, tasks in segments:
            if _stopped():
                break
            log(f"======== {adapter.title}：{len(tasks)} 个任务 ========")
            fn = (lambda i, n, st, info=None, o=offset: status_fn(o + i, total, st, info)) if status_fn else None
            results += run_tasks(tasks, adapter, username, password, fn, session=session)
            offset += len(tasks)
    except StopRequested:
        log("用户已停止下载")
    finally:
        force_quit()
    return results
