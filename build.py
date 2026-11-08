# -*- coding: utf-8 -*-
"""
打包脚本：把统一界面打包成 Windows exe（PyInstaller）
================================================================
    python build.py              # 单文件 exe（dist\\影子科技数据下载器.exe）
    python build.py --onedir     # 文件夹版：启动更快、杀毒软件误报更少（dist\\影子科技数据下载器\\）
    python build.py --console    # 带控制台窗口（排查问题时用，能看到全部日志）
    python build.py --no-install # 跳过 pip 安装依赖

需要的文件（与本脚本放在同一目录）：
    app_gui.py  yz_core.py  task_report.py  task_transfer.py  task_sowfile.py  task_genetic.py  task_ebv.py  task_pedigree.py
    schedule_plan.py  ui_icons.py
    requirements.txt
可选：icon.ico（程序图标）
"""

import os
import sys
import shutil
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
NAME = "影子科技数据下载器"
ENTRY = "app_gui.py"
MODULES = ["yz_core", "task_report", "task_transfer", "task_sowfile", "task_genetic", "task_ebv", "task_pedigree", "schedule_plan", "ui_icons"]

ENV_EXAMPLE = """# 把本文件改名为 .env（与 exe 放在同一目录），填入账号密码后，启动时会自动填入界面
YINGZI_USERNAME=你的账号
YINGZI_PASSWORD=你的密码
# 可选：新电脑第一次打开时的默认保存目录（之后在界面里改的会被记住）
# YINGZI_DOWNLOAD_DIR=D:\\影子科技数据下载
"""


def run(cmd):
    print(">>", " ".join(cmd), flush=True)
    subprocess.check_call(cmd, cwd=HERE)


def main():
    if sys.version_info < (3, 8):
        sys.exit("需要 Python 3.8 或更高版本")
    missing = [f for f in [ENTRY] + [m + ".py" for m in MODULES] if not os.path.exists(os.path.join(HERE, f))]
    if missing:
        sys.exit(f"缺少文件：{', '.join(missing)}")

    onedir = "--onedir" in sys.argv
    console = "--console" in sys.argv

    if "--no-install" not in sys.argv:
        run([sys.executable, "-m", "pip", "install", "-U", "-r", "requirements.txt"])

    # 打包前先确认能正常导入（语法错误、缺依赖在这里就能发现）
    run([sys.executable, "-c", "import " + ", ".join(MODULES) + "; print('导入检查通过')"])

    for d in ("build", "dist"):
        shutil.rmtree(os.path.join(HERE, d, NAME), ignore_errors=True)

    cmd = [sys.executable, "-m", "PyInstaller", ENTRY,
           "--name", NAME, "--noconfirm", "--clean",
           "--onedir" if onedir else "--onefile",
           "--paths", HERE,
           # selenium 4 自带 selenium-manager（自动匹配 chromedriver），必须连同数据文件一起打包
           "--collect-all", "selenium",
           "--collect-submodules", "webdriver_manager",
           "--hidden-import", "openpyxl",
           "--hidden-import", "xlrd",
           "--exclude-module", "matplotlib",
           "--exclude-module", "numpy",
           "--exclude-module", "pandas"]
    for m in MODULES:
        cmd += ["--hidden-import", m]
    if not console:
        cmd.append("--windowed")
    icon = os.path.join(HERE, "icon.ico")
    if os.path.exists(icon):
        cmd += ["--icon", icon]
    run(cmd)

    out_dir = os.path.join(HERE, "dist", NAME) if onedir else os.path.join(HERE, "dist")
    with open(os.path.join(out_dir, ".env.example"), "w", encoding="utf-8") as f:
        f.write(ENV_EXAMPLE)
    env = os.path.join(HERE, ".env")
    if os.path.exists(env) and not os.path.exists(os.path.join(out_dir, ".env")):
        shutil.copy(env, out_dir)
        print("已把 .env 复制到输出目录（发给别人前记得删掉）")

    exe = os.path.join(out_dir, NAME + ".exe")
    print("\n" + "=" * 60)
    print(f"打包完成：{exe}")
    print("首次运行建议：取消「后台运行」，登录一次（如有验证码手动完成）并勾选「记住登录状态」。")
    print("程序目录下会生成：browser_profile（登录缓存）、debug（出错截图）、设置文件和日志。")


if __name__ == "__main__":
    main()
