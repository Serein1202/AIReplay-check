#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Replay 核查一体化脚本（图形界面版）
=================================================================
一次运行完成「后台抓取 → CSV 导出 → 筛选/排序/去重 → xlsx 输出」全流程，
带可视化操作界面（customtkinter）。

界面用法:
  1. 打开窗口后在「工作台」选日期（默认「昨天」，可点快捷日期或自绘月历），点「开始核查」
  2. 四段阶段指示（抓取 → 导出 → 处理 → 复制）配细进度条，三块统计实时更新；日志可展开
  3. 完成后就地出现「已复制 C1:L11」提示条，下方的预览表就是将被复制的内容
  4. 左侧导航可切到「历史记录」（对已有产物再次复制 / 打开所在目录）与「设置」
     （筛选、排序、去重、剪贴板区域、完成后是否弹窗、界面动效、深浅色主题）

输出与剪贴板:
  - xlsx 生成后不会自动打开文件（也不弹 Excel），只在结果卡片里展示路径
  - 生成后立即把 xlsx 的 CLIPBOARD_RANGE 区域（默认 C1:L11）复制到剪贴板，
    同时写「制表符文本 + HTML 表格」两种格式：
      粘到 Excel / WPS → 按单元格铺开；粘到微信 / 飞书 / Word → 直接变成表格
  - 区域数据不足（当天行数少于 11）时自动裁掉尾部空行 / 空列，不会粘出一串空制表符
  - 剪贴板被其他程序占用导致失败时只记一条警告，报表本身不受影响
  - 结果卡片提供「再次复制 C1:L11」，无需重跑抓取
  - 关闭：设置页关掉「输出后自动复制」（写进 .aireplay_config.json 的 clipboard_enabled）

命令行用法:
  python AIReplay-check.py 20260924         # 静默模式（自动化 / 定时任务）
  python AIReplay-check.py --date 20260924  # 同上
  python AIReplay-check.py --selfcheck      # 界面自检（启动后自动退出）
  python AIReplay-check.py --check-update   # 只检测更新并打印结果（不启动界面）
  python AIReplay-check.py --version        # 打印当前版本号

软件授权（License Guard）:
  - 由自建授权服务签发令牌，客户端仅做验签，校验全程在后台进行
  - 全程静默：任何情况下都不能让人看出软件带授权 ——
      · 界面（日志区 / 状态栏 / 按钮）不显示任何授权信息
      · 任何情况下都不弹授权对话框（即使后台把「到期时的行为」设成「弹提示后退出」）
      · 命令行同样零输出：没有 --license 之类的子命令，校验失败不打印任何字样
      · 诊断信息默认全部丢弃；排障时设 AIREPLAY_LICENSE_TRACE=<文件路径> 才写盘
  - 授权码不写进仓库：仓库里 LICENSE_KEY 为空，正式构建时由 CI 从 GitHub Secrets
    注入到脚本里再打包（这样公开仓库和 git 历史里都查不到授权码）
    （运行时的覆盖优先级：环境变量 LICENSE_KEY > 程序同目录 license.key > 脚本内置 LICENSE_KEY）
  - 其余策略完全跟随后台设置，客户端不自行揣测：
      心跳间隔、断网容忍天数、到期后宽限、设备数上限、令牌有效期
      全部由后台下发并原样执行（后台的「到期时的行为」仅保留退出码效果）
  - 每次启动都联网校验（LICENSE_ONLINE_CHECK_SEC = 0），后台改吊销 / 过期立刻生效；
    只有连不上服务器时才按后台的「断网容忍天数」离线放行
  - 命令行模式与降级弹窗模式同样要过校验，不存在绕过路径
  - 失败只体现为「进程退出码」，不产生任何文字 / 对话框

登录态说明:
  获取顺序（优先级由高到低）:
    1. 环境变量 AMIS_TOKEN
    2. 本地 Token 缓存（脚本 / exe 同目录下的 .amis_token.json，未过期时直接复用）
    3. 本机 Chrome / Edge 中 amis.ssjj.cn 的有效 Token（AMIS-USER-TOKEN）
    4. 以上均不可用（未找到或已过期）时，先尝试用本机加密保存的账号密码自动登录
       换取新 Token（全程不弹框）；自动登录不可用（未保存凭据 / 凭据被服务端拒绝 /
       网络或服务异常）时才弹出登录框手动登录
  本地 Token 缓存:
    - 登录成功（含自动弹框登录）后统一写入缓存，内容仅 token / account / expire / saved_at
    - 缓存命中且未过期 → 直接使用并在日志中提示来源与有效期，不再要求重新登录
    - 使用缓存 / 浏览器 Token 抓取时若被服务端判定失效（HTTP 401/403 或返回鉴权、
      登录失效类错误码 / 消息），自动删除缓存文件并弹出登录框换取新 Token，
      登录成功后覆盖写入缓存并继续本次抓取，无需用户重跑
    - 缓存文件损坏或解析失败时按无缓存处理，并删除该文件
    - 密码仅在内存中使用，绝不写入缓存、日志或代码
  本机加密凭据（.amis_credentials.bin）:
    - 登录框勾选「记住账号密码（本机加密保存）」且登录成功后，用 Windows DPAPI
      （CryptProtectData，ctypes 调用，不新增 pywin32 依赖）将
      {account, password, ldap, saved_at} 加密写入脚本 / exe 同目录下的
      .amis_credentials.bin，密码不以明文落盘
    - Token 缺失 / 过期 / 被服务端判定失效时，优先用已保存凭据自动登录换取新
      Token；凭据被服务端拒绝（如账号密码错误）时自动删除本地凭据，再回退到
      登录框手动登录
    - 登录框取消勾选「记住账号密码」后登录成功，会删除已保存的本地凭据；
      凭据文件损坏或无法解密时按未保存处理并自动删除
  登录框:
    - 框内展示登录网址 http://amis.ssjj.cn/，含账号、密码（掩码显示）、
      「LDAP 登录」复选项（默认不勾选，使用 LDAP 账号时勾选）与
      「记住账号密码（本机加密保存）」复选项（默认勾选）
    - 调用 POST http://amis.ssjj.cn/auth/signIn（Header X-PROJECT-ID: 1，
      体 {"account","password","ldap"}）换取 Token
    - 登录失败会提示服务端返回的 msg 并允许重试（最多 3 次）；全部失败或取消则中止本次抓取
    - 有界面（customtkinter）时用深色对话框（CTkToplevel）；无界面时退回原生 tkinter 对话框
    - 账号密码仅在本机内存中使用，不写入磁盘、日志或代码；脚本内不含任何硬编码凭据

在线更新:
  - 打开程序后自动检测最新版本；发现更高版本时弹窗提示，可一键下载并重启到新版本
  - 新版本下载到 exe 同目录，文件名形如 AIReplay-<版本号>.exe；
    重启后由新版本自动清理旧版本文件，目录里不会堆积历史 exe
  - 检测与下载都是「直连优先，失败自动切换加速线路」：
    见 UPDATE_API_MIRRORS / UPDATE_ASSET_MIRRORS，国内网络下能显著减少「下载失败」；
    下载失败会自动换下一条线路重试，全部线路都不通才报错
  - 界面（更新说明、失败提示、日志区）一律不出现仓库地址，
    也不引导用户去外部页面手动下载；网络不可达时静默跳过，不影响正常使用
  - 关闭自动检测: 设置环境变量 AIREPLAY_NO_UPDATE=1，或把 UPDATE_AUTO_CHECK 改为 False
  - 版本号来源: 打包时由 CI 注入 APP_VERSION（见 build-release.yml）；
    以源码方式运行时显示占位版本，且不提供在线更新

依赖:
  pip install -r requirements.txt

打包 exe:
  pyinstaller --noconsole --onefile --clean --collect-all customtkinter \
    --icon "assets/app.ico" --add-data "assets/app.ico;assets" \
    --name "AIReplay" AIReplay-check.py

自动构建与发布:
  推送到 main 分支 → GitHub Actions 自动打包 exe 并发布到仓库 Releases，
  发布文件命名为 AIReplay-<版本号>.exe（如 AIReplay-v1.0.1.exe）
  配置见 .github/workflows/build-release.yml；版本号取历史最高 tag 递增一位
  （patch 只跑 0~9，到 9 进 minor），由 CI 自动打 tag
"""

import base64
import calendar
import csv
import ctypes
from ctypes import wintypes
import json
import os
import queue
import re
import shutil
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from datetime import date, datetime, timedelta

# ====================== 可配置项 ======================
# 后台接口
HOST = "http://amis.ssjj.cn"
API_PATH = "/gameConfigs/listAiHackCheckReport"
PROJECT_ID = "28"
STATUS_DONE = 3          # 3 = 已核查
PER_PAGE = 1000          # 单页拉取条数（实测 1000 可一次取全）
HTTP_TIMEOUT = 30        # 单次请求超时（秒）

# 登录接口（自动获取 Token 失败 / Token 过期时的兜底：手动登录换取新 Token）
LOGIN_PAGE = HOST + "/"          # 登录页展示地址：http://amis.ssjj.cn/
LOGIN_API = HOST + "/auth/signIn"
LOGIN_PROJECT_ID = "1"           # 登录接口专用项目号，固定 "1"（用 PROJECT_ID=28 会返回 400）
LOGIN_MAX_RETRY = 3              # 手动登录最大重试次数

# 本地 Token 缓存：保存于脚本 / exe 同目录，避免每次抓取都重新登录
TOKEN_CACHE_NAME = ".amis_token.json"

# 本机加密凭据：保存于脚本 / exe 同目录（Windows DPAPI 加密，用于 Token 失效时自动登录）
CREDENTIALS_FILE_NAME = ".amis_credentials.bin"

# 输出命名（YYYYMMDD 由输入日期填充）
CSV_NAME_TPL = "replay已核查-{day}.csv"
XLSX_NAME_TPL = "replay核查-{day}.xlsx"

# 报表生成后自动复制到剪贴板的区域（Excel 的 A1 表示法，含首尾单元格）。
# 复制内容同时提供「制表符分隔文本」与「HTML 表格」两种格式：
#   · 粘到 Excel / WPS  → 按单元格铺开，保留行列
#   · 粘到微信 / 飞书 / Word → 直接变成表格
# 设为空串 "" 或把 CLIPBOARD_ENABLED 改 False 即可关闭该行为（关闭后不产生任何副作用）。
CLIPBOARD_ENABLED = True
CLIPBOARD_RANGE = "C1:L11"

# 目标文件已存在时，是否先重命名保留旧文件（True=防覆盖，False=直接覆盖）
PROTECT_EXISTING = True

# 筛选条件：列名含以下关键字（不区分大小写），精确值匹配
FILTER_COLUMNS = {
    "封禁状态": "未封禁",
    "备注":     "无异常",
}
# 排序：列名含以下关键字（不区分大小写），默认降序
SORT_COLUMN_KEYWORD = "分数"
# 去重依据：列名含以下关键字（不区分大小写），保留第一次出现的行
DEDUP_COLUMN_KEYWORD = "角色ID"
# 想设为"文本"的列：按表头关键字匹配（不区分大小写），命中即转换
TEXT_COLUMN_KEYWORDS = ["id", "房间"]
# 想设为"日期时间"的列：按表头关键字匹配（不区分大小写），命中即转换
DATETIME_COLUMN_KEYWORDS = ["开始时间", "结束时间", "上报时间"]
# 也可直接写死列字母（优先级最高），例如 ["A", "C"]
TEXT_COLUMN_LETTERS = []

# CSV 表头（与页面表格列一一对应）
CSV_HEADERS = [
    "#", "处理状态", "角色名", "角色ID", "无端ID", "房间号", "分数",
    "开始时间", "结束时间", "战斗耗时", "上报时间", "处理人", "处理时间", "封禁状态", "备注",
]
# 处理状态码 → 文案
STATUS_TEXT = {3: "已核查"}

# ====================== Sentry 监控配置 ======================
SENTRY_DSN = "https://1164181e93905c40acb94ed3c27679c1@o4511234196570112.ingest.us.sentry.io/4512146535940096"
SENTRY_TRACES_SAMPLE_RATE = 1.0
# ============================================================

APP_TITLE = "AIReplay 核查工具"
FONT_UI = "Microsoft YaHei UI"

# ====================== 版本与在线更新 ======================
# 当前版本号：仓库源码里保留占位值，打包时由 GitHub Actions 注入该次发布的真实版本号
# （见 .github/workflows/build-release.yml 的「注入版本号」步骤）
#
# 发布方式：push 到 main 就自动发版，版本号由 CI 自动计算。
#     git push            # 就这一步
#   CI 取历史最高 tag 递增一位（如 v1.0.9 → v1.1.0），自动打 tag 并发布 Release。
#   号完全由历史 tag 推算，天然单调递增；CI 还会做一次倒序保护校验。
#   何时需要手动打 tag：只在「想指定某个特定版本号」时才需要，平时不用管。
APP_VERSION = "0.0.0"
APP_VERSION_PLACEHOLDER = "0.0.0"

# 在线更新：查询该仓库的最新 Release，发现更高版本时提示并可自动下载、重启到新版本
GITHUB_REPO = "Serein1202/AIReplay-check"
UPDATE_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
UPDATE_TIMEOUT = 15                # 检测 / 下载超时（秒）
UPDATE_DL_TIMEOUT = 25             # 下载单条线路的超时（秒）。下载比检测宽松些
UPDATE_DL_ATTEMPTS = 3             # 同一条线路内最多尝试几次（每次从断点续传）
UPDATE_ASSET_PREFIX = "AIReplay-"  # Release 中 exe 资产的文件名前缀
UPDATE_AUTO_CHECK = True           # 启动时自动检测更新（环境变量 AIREPLAY_NO_UPDATE=1 可临时关闭）
UPDATE_MARKER = ".aireplay_update.json"  # 记录上次在线更新后待清理的旧文件

# 下载加速线路（前缀式镜像：最终地址 = 前缀 + 原始地址）。
# 国内直连不稳定时，自动依次切换，避免「下载失败」。
#   · 检测更新（api）与下载（releases）都能走：gh-proxy.org
#   · 只能代理文件下载：ghproxy.net
# 线路只影响「去哪儿取包」，界面上不会出现任何线路地址。
UPDATE_API_MIRRORS = (
    "https://gh-proxy.org/",
)
UPDATE_ASSET_MIRRORS = (
    "https://gh-proxy.org/",
    "https://ghproxy.net/",
)
# 下载完整性下限：正常产物约 35 MB，小于该值视为没下到真东西
UPDATE_MIN_BYTES = 1024 * 1024

# ====================== 授权验证（License Guard）======================
# 由自建 Cloudflare Worker（license-guard）签发授权令牌，客户端只做验签：
#   · 到期时间写在服务端签发的令牌里，本地改不了
#   · 系统时间被回拨会被检测（CLOCK_TAMPERED）
#   · 断网超过服务端设定的「离线容忍天数」后必须联网校验，否则失效
# 后台（管理页面）可随时调整到期时间、设备数、离线容忍天数、功能开关等策略。
#
# 校验方式：启动后立即在后台线程静默完成，界面不展示任何授权信息，
#          也不改动按钮 / 状态栏 / 日志区；未通过时程序自行退出。
LICENSE_ENABLED = True                                    # False = 跳过授权校验（仅供本地调试）
LICENSE_SERVER = "https://license-guard.xhd19981202.workers.dev"
LICENSE_APP_ID = "583729164"                              # 后台创建软件时填的 app_id
LICENSE_PUBLIC_KEY = "LjZLPnhcuJ5qJEtNi1axHbc54lvk-wfN5wcCTfZzIEk"  # Ed25519 公钥(b64url)
LICENSE_SIGN_ALG = "ed25519"                              # ed25519 | es256
# 内置授权码。**仓库里刻意留空** —— 正式构建时由 CI 从 GitHub Secrets 注入
# （见 .github/workflows/build-release.yml 的「注入授权码到脚本」步骤）。
# 这样公开仓库及 git 历史里都查不到授权码，换码也只需改 Secret 重跑构建。
# 本地运行 / 本地打包时的回退顺序（见 resolve_license_key）：
#   1. 环境变量 LICENSE_KEY
#   2. 程序同目录（或当前工作目录）的 license.key   ← 本地开发推荐，已在 .gitignore
# 三项都为空时程序不会启动（拿不到授权码）。
LICENSE_KEY = ""
LICENSE_FILE_NAME = "license.key"                         # 备用授权码文件（内置授权码有效时可不存在）
LICENSE_TIMEOUT = 8                                       # 单次校验请求超时（秒）
# 运行期复查间隔（秒）—— 决定「程序一直开着」时多久能感知到后台的
# 吊销 / 改期 / 设过期。0 = 不指定，**完全跟随后端下发的「心跳间隔」**
# （后端当前是 21600 秒 = 6 小时，即运行中最长 6 小时感知到吊销）。
# 若想更快生效，把它调小（例如 120），并同时把下面的上限也设成同样的值。
LICENSE_WATCHDOG_SEC = 0
# 复查间隔上限（秒）。0 = 不封顶，完全跟随后端心跳间隔（当前即 6 小时一次）。
# 注意：SDK 只把 watchdog_interval_sec 当「第一次复查」的等待时间，之后的间隔会
# 重新按后端心跳计算；所以想缩短复查间隔，这两项要一起改才有效。
LICENSE_HEARTBEAT_CAP_SEC = 0
# 缓存令牌「免联网直接放行」的宽限时长（秒）：
#   0    = 每次校验都联网，后台吊销 / 改期 / 设过期立即生效（推荐，默认）
#   >0   = 该时长内复用缓存不再联网，仅在超时后重新联网
# 注意：无论取何值，只要联不上服务器，仍会按后端下发的「断网容忍天数」放行
LICENSE_ONLINE_CHECK_SEC = 0
# 以下两个只是「后端够不着时」的兜底默认值：
# 正常情况下退出行为（闪退 / 弹提示）与退出码完全由后台
# 「到期时的行为」和 exit_code 决定（拒绝响应里随策略一起下发）
LICENSE_FAIL_MODE = "hard"                                # hard = 静默闪退（兜底）
LICENSE_EXIT_CODE = 1                                     # 退出码（兜底）

# ---- 以下是内部实现，一般无需修改 ----

try:
    if SENTRY_DSN:
        import sentry_sdk
        sentry_sdk.init(
            dsn=SENTRY_DSN,
            traces_sample_rate=SENTRY_TRACES_SAMPLE_RATE,
            environment="production",
            release=os.path.basename(__file__),
        )
except Exception as _e:  # Sentry 不可用不影响主流程
    try:
        print(f"[警告] Sentry 初始化失败（不影响运行）: {_e}", flush=True)
    except Exception:
        pass

def _license_trace(msg):
    """授权诊断输出 —— **默认什么都不输出**。

    用户硬性要求：任何情况下都不能让人看出这个软件设置了授权。
    因此这里既不 print、也不弹窗、也不写程序目录，默认完全是空操作。

    需要排障时，提前设一个环境变量即可把诊断写进文件（注意：别用 stdout，
    否则从终端启动时会被终端 / 调度器日志捕获）：

        set AIREPLAY_LICENSE_TRACE=D:\tmp\aireplay_license.log
    """
    path = os.environ.get("AIREPLAY_LICENSE_TRACE")
    if not path:
        return
    try:
        with open(path, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {msg}\n")
    except Exception:
        pass


# ---- 授权 SDK（随程序一起分发，零第三方依赖；缺失时降级为「不校验」）----
try:
    import license_guard
    _LICENSE_SDK_OK = True
except Exception as _le:
    license_guard = None
    _LICENSE_SDK_OK = False
    try:
        _license_trace(f"授权模块 license_guard 不可用: {_le}")
    except Exception:
        pass


def report_exception(exc, **tags):
    """把异常上报到 Sentry（未配置 DSN 或未安装 SDK 时静默跳过）"""
    try:
        if not SENTRY_DSN:
            return
        import sentry_sdk
        if tags:
            sentry_sdk.set_tags(tags)
        sentry_sdk.capture_exception(exc)
    except Exception:
        pass


# ---------------------------------------------------------------- 基础工具
def script_dir():
    """脚本所在目录（PyInstaller 打包成 exe 时取 exe 所在目录）"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def app_icon_path():
    """窗口图标 assets/app.ico 的路径；找不到返回 None（没图标也能正常跑）。

    打包成 onefile 后，--add-data 加进来的文件被解到 sys._MEIPASS 下，
    所以要先看那里，再回落到脚本目录（源码直接运行时用）。
    """
    roots = [getattr(sys, "_MEIPASS", None), script_dir()]
    for root in roots:
        if not root:
            continue
        p = os.path.join(root, "assets", "app.ico")
        if os.path.isfile(p):
            return p
    return None


def _console(msg):
    """控制台输出；exe 无控制台（--noconsole）时静默忽略"""
    try:
        print(msg, flush=True)
    except Exception:
        pass


def _exit_now(code=0):
    """立即结束进程，跳过解释器收尾阶段，保证关窗口后进程马上消失。

    onefile 打包的程序在退出时，解释器收尾（atexit 回调、非守护线程等待、
    大模块析构）加上杀毒软件对解包目录的扫描，会让进程在窗口关闭后仍残留
    很久。这里只做必要的收尾（上报 flush、标准输出 flush），然后直接结束。
    """
    try:
        n = int(code or 0)
    except Exception:
        n = 0
    # Sentry 若已加载，给它一次短暂的机会把事件发出去
    _sentry = sys.modules.get("sentry_sdk")
    if _sentry is not None:
        try:
            _sentry.flush(timeout=1.5)
        except Exception:
            pass
    for _stream in (sys.stdout, sys.stderr):
        try:
            _stream.flush()
        except Exception:
            pass
    os._exit(n)


# 界面日志接收器：GUI 启动时被替换为向界面推送的回调，其余场景为 None
_LOG_SINK = None


def log(msg):
    """统一日志出口：控制台 + 界面日志区"""
    _console(msg)
    sink = _LOG_SINK
    if sink is not None:
        try:
            sink(f"{datetime.now():%H:%M:%S}  {msg}")
        except Exception:
            pass


def popup(title, text, kind="info"):
    """原生弹窗提示（仅在无 GUI 依赖的降级路径中使用）"""
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        getattr(messagebox, kind)(title, text)
        root.destroy()
    except Exception:
        log(f"[{title}] {text}")


def ask_date(default_day):
    """原生弹窗输入抓取日期，返回 YYYYMMDD 字符串；取消返回 None"""
    import tkinter as tk
    from tkinter import simpledialog, messagebox
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    while True:
        val = simpledialog.askstring(
            "抓取日期",
            "请输入抓取日期（格式：20260924）",
            initialvalue=default_day,
            parent=root,
        )
        if val is None:
            root.destroy()
            return None
        cleaned = re.sub(r"[^0-9]", "", val)
        try:
            datetime.strptime(cleaned, "%Y%m%d")
            root.destroy()
            return cleaned
        except ValueError:
            messagebox.showwarning(
                "格式错误", "日期格式不正确，请按 20260924 的格式重新输入。", parent=root
            )


def backup_if_exists(path):
    """目标文件已存在时先重命名保留，避免覆盖既有产物"""
    if not (PROTECT_EXISTING and os.path.exists(path)):
        return
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base, ext = os.path.splitext(path)
    new_path = f"{base}（旧-{stamp}）{ext}"
    try:
        os.replace(path, new_path)
        log(f"[提示] 目标文件已存在，已重命名保留旧文件: {new_path}")
    except Exception as e:
        log(f"[警告] 旧文件重命名失败（将继续写入）: {e}")


def open_folder(path):
    """在资源管理器中打开目录。

    [重要] Windows 的 ShellExecute 对「不带扩展名」的路径会先按 PATHEXT 规则
    逐项尝试补扩展名（.exe/.com/.bat/...），全部落空后才按目录处理。
    所以当输出目录 ...\\replay核查\\replay核查 的同级存在 replay核查.exe 时，
    os.startfile(path) 会被解析成「启动 replay核查.exe」——现象就是
    「点打开输出目录却又新运行了一个 exe」，而不是打开目录。
    修复: Windows 下改用 explorer 打开；并以追加尾分隔符作为兜底。
    """
    import subprocess

    path = os.path.normpath(path)
    os.makedirs(path, exist_ok=True)

    if hasattr(os, "startfile"):  # Windows
        try:
            subprocess.Popen(["explorer", path])
        except OSError:
            # 兜底: 尾分隔符强制按目录处理，避免被同名 .exe 抢走
            os.startfile(path + os.sep)
        log(f"[提示] 已在资源管理器中打开目录: {path}")
    else:
        subprocess.Popen(["xdg-open", path])
        log(f"[提示] 已打开目录: {path}")
    return path


# ==================================================================
#                      版本与在线更新（GitHub Releases）
# ==================================================================
def current_exe_path():
    """当前可执行文件路径；以源码方式（python xxx.py）运行时返回 None"""
    if getattr(sys, "frozen", False):
        return os.path.abspath(sys.executable)
    return None


def current_app_version():
    """当前版本号（不含前导 v）。

    优先级：
      1. APP_VERSION —— 打包时由 CI 注入，代表这个 exe 的真实版本
      2. exe 文件名里的版本（AIReplay-v1.0.1.exe → 1.0.1），用于本地打包 / 手工改名场景
    """
    v = (APP_VERSION or "").strip().lstrip("vV")
    if v and v != APP_VERSION_PLACEHOLDER:
        return v
    if getattr(sys, "frozen", False):
        m = re.search(r"AIReplay[-_]?v?(\d+(?:\.\d+)*)",
                      os.path.basename(sys.executable), re.I)
        if m:
            return m.group(1)
    return v or APP_VERSION_PLACEHOLDER


def _version_key(value):
    """把版本号转成可比较的元组：1.0.10 > 1.0.9"""
    nums = re.findall(r"\d+", str(value))
    return tuple(int(n) for n in nums[:6]) if nums else (0,)


def version_is_newer(latest, current):
    """latest 是否比 current 更新"""
    a, b = _version_key(latest), _version_key(current)
    n = max(len(a), len(b))
    return a + (0,) * (n - len(a)) > b + (0,) * (n - len(b))


def is_dev_build():
    """当前是不是「本地开发/自测构建」——版本号识别不出来（仍为占位值）时算开发版。

    为什么要单独判断：源码方式运行、或本地 PyInstaller 打出来的包（如 AIReplay-local.exe）
    名称里不带版本号，current_app_version() 只能回落到 0.0.0。
    这种包拿去和线上 Release 比，永远判「有新版本」，每次启动都弹更新框，
    纯属噪音。因此开发构建直接跳过自动更新检查。
    """
    v = (APP_VERSION or "").strip().lstrip("vV")
    if v and v != APP_VERSION_PLACEHOLDER:
        return False
    if getattr(sys, "frozen", False):
        if re.search(r"AIReplay[-_]?v?(\d+(?:\.\d+)*)",
                     os.path.basename(sys.executable), re.I):
            return False
    return True


# ====================== 更新文案清洗 ======================
# 更新说明（取自发布正文）与更新失败提示里**不允许出现仓库地址**，
# 也不允许引导用户去仓库手动下载。这里做统一兜底：即便正文里被写进了链接、
# 托管平台名或 owner/repo 形式的标识，界面上也只会显示去掉这些痕迹后的文字。
#
# 另一个职责：更新说明区是**纯文本框，不会渲染 Markdown**，所以正文里的
# `#` 标题、`**加粗**`、反引号、表格竖线等记号会原样显示成一堆符号。
# 这里先用 md_to_plain() 把记号转成普通文字，再做地址清洗。
_URL_RE = re.compile(r"(?:https?://|www\.)[^\s<>()\[\]{}\"'，。；、）】]+", re.I)
_HOST_RE = re.compile(r"\b(?:[\w.\-]*\.)?github(?:usercontent)?\.(?:com|io)\b", re.I)
_SLUG_RE = re.compile(r"\b[\w.\-]+/[A-Za-z0-9._\-]*AIReplay[A-Za-z0-9._\-]*\b")
# 删掉域名 / 仓库标识后可能残留的路径片段（如 "/releases"、"//releases"）
_PATH_RE = re.compile(r"(?<![\w.\-])/[\w.\-/]*")
_BRAND_RE = re.compile(r"\bGitHub\b", re.I)
# 加速线路的域名同样只在程序内部使用，界面上一个字都不能出现
_MIRROR_HOSTS = sorted({
    urllib.parse.urlsplit(_p).hostname
    for _p in (tuple(UPDATE_API_MIRRORS) + tuple(UPDATE_ASSET_MIRRORS))
    if urllib.parse.urlsplit(_p).hostname
})
_MIRROR_RE = (
    re.compile(r"\b(?:[\w-]+\.)*(?:" + "|".join(re.escape(_h) for _h in _MIRROR_HOSTS) + r")\b", re.I)
    if _MIRROR_HOSTS else None)
_GUIDE_RE = re.compile(
    r"(?:Full\s*Changelog|Changelog|详见|参见|请见|可见|此链接|该链接|链接|"
    r"手动下载|自行下载|重新下载|下载地址|请前往|请访问|请到|前往|访问|点击|官网)",
    re.I)
_TRIM_CHARS = " \t\u3000，,、;；:：-—–·•/\\()（）[]【】<>\"'|`*_#~"

# 构建细节行：对用户没有意义，整行丢弃（正文若仍是旧版自动生成的元信息，
# 也不会把「GitHub Actions / CI / tag」这类字眼露给用户）。
_META_RE = re.compile(
    r"(?:GitHub\s*Actions|\bActions\b|由\s*CI\b|\bCI\b|自动构建|自动递增|"
    r"历史\s*tag|Full\s*Changelog|push\s*到|Release\s*由|^提交\s*[:：])",
    re.I)

# ---------- Markdown 记号（说明区不渲染 Markdown，必须先去掉这些符号） ----------
_FENCE_RE = re.compile(r"^\s*(?:```|~~~)")
_HEAD_RE = re.compile(r"^\s{0,3}#{1,6}\s*")
_QUOTE_RE = re.compile(r"^\s{0,3}>\s?")
_RULE_RE = re.compile(r"^\s*(?:[-*_]\s*){3,}$")
_TBL_SEP_RE = re.compile(r"^\s*\|?\s*:?-{2,}:?\s*(?:\|\s*:?-{2,}:?\s*)*\|?\s*$")
_BULLET_RE = re.compile(r"^\s{0,3}(?:[-*+]|\u2022)\s+")
_LINK_RE = re.compile(r"!?\[([^\]\n]*)\]\([^)\n]*\)")
_CODE_RE = re.compile(r"`+([^`\n]*)`+")
_BOLD_RE = re.compile(r"(\*\*|__)(?=\S)(.+?)(?<=\S)\1")
_ITAL_RE = re.compile(r"(?<![\w*])([*_])(?=\S)([^*_\n]+?)(?<=\S)\1(?![\w*])")
_ESC_RE = re.compile(r"\\([\\`*_{}\[\]()#+\-.!|>])")

# 表格表头（「项目 / 内容」这类只有栏名没有信息的行）→ 整行丢弃
_TBL_HEAD_WORDS = {"项目", "内容", "名称", "说明", "字段", "值", "项",
                   "item", "value", "name", "key", "title", "field"}


def _md_table_row(line):
    """把 `| 项目 | 内容 |` 一行转成 `项目：内容` 这样的普通文字。"""
    cells = [c.strip() for c in line.strip().strip("|").split("|")]
    cells = [c for c in cells if c]
    if not cells:
        return ""
    if len(cells) == 2:
        if cells[0].lower() in _TBL_HEAD_WORDS and cells[1].lower() in _TBL_HEAD_WORDS:
            return ""                        # 纯表头（如「项目 / 内容」），没有信息量
        return f"{cells[0]}：{cells[1]}"
    return " · ".join(cells)


def md_to_plain(text):
    """把 Markdown 记号转成普通文字（只去记号，不做任何渲染）。

    说明区是纯文本框，若不处理就会把 `## 标题`、`**加粗**`、表格竖线原样显示出来。
    """
    if not text:
        return ""
    out = []
    in_fence = False
    for raw in str(text).replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = raw.rstrip()
        if _FENCE_RE.match(line):
            in_fence = not in_fence          # 围栏行本身丢掉，内容保留
            continue
        stripped = line.strip()
        if not in_fence:
            if _RULE_RE.match(stripped) or _TBL_SEP_RE.match(line):
                continue                     # 分隔线 / 表格分隔行：整行丢弃
            if not stripped:
                out.append("")               # 保留空行（段落间距）
                continue
            line = _QUOTE_RE.sub("", line)
            line = _HEAD_RE.sub("", line)
            line = _ESC_RE.sub(r"\1", line)
            line = _LINK_RE.sub(r"\1", line)
            line = _CODE_RE.sub(r"\1", line)
            line = _BOLD_RE.sub(r"\2", line)
            line = _ITAL_RE.sub(r"\2", line)
            body = line.lstrip()
            if body.startswith("|") and "|" in body[1:]:
                line = _md_table_row(line)
            else:
                line = _BULLET_RE.sub("\u00b7 ", line)
        out.append(line)
    return "\n".join(out)


def scrub_update_text(text):
    """把展示用文本整理成「可以直接给用户看」的纯文本。

    做两件事：
      1. md_to_plain()：去掉 Markdown 记号（说明区是纯文本框，不渲染 Markdown）；
      2. 去掉仓库痕迹（地址 / 平台名 / owner-repo 标识 / 下载引导）与构建细节行。

    只作用于「给用户看的文本」，程序内部仍按原样使用原始下载地址。
    整行若只有一条链接、或去链接后只剩零星引导词，则该行直接丢弃，
    避免界面上留下「详见」这种读不通的残句。
    """
    if not text:
        return ""
    out = []
    blank = False                              # 上一行是空行 → 下一条正文前补一个空行
    for raw in md_to_plain(text).splitlines():
        if _META_RE.search(raw):
            continue                           # 构建细节行：整行不展示
        hit = bool(_URL_RE.search(raw) or _HOST_RE.search(raw) or _SLUG_RE.search(raw)
                   or (_MIRROR_RE and _MIRROR_RE.search(raw)))
        line = _URL_RE.sub("", raw)
        line = _HOST_RE.sub("", line)
        if _MIRROR_RE:
            line = _MIRROR_RE.sub("", line)
        line = _SLUG_RE.sub("", line)
        line = _BRAND_RE.sub("更新服务", line)
        if hit:
            line = _PATH_RE.sub("", line)      # 清掉剩下的 /releases 之类路径残片
            line = _GUIDE_RE.sub("", line)
            line = line.strip().strip(_TRIM_CHARS).strip()
        else:
            line = line.strip()
        # 含过链接的行，若剥掉链接后只剩零星几个字（如「详见」），整行丢弃
        if hit and len(re.sub(r"[\W_]+", "", line)) <= 4:
            line = ""
        if not line:
            blank = True                       # 空行只记标志，避免连续空行
            continue
        if blank and out:
            out.append("")                     # 段落之间保留一个空行，读起来不挤
        blank = False
        out.append(line)
    while out and not out[-1]:
        out.pop()
    return "\n".join(out)


def _fetch_release_json():
    """取最新 Release 的 JSON：直连优先，失败后依次走加速线路。

    返回 (data, err_msg)；err_msg 为 None 表示成功。错误文案不含任何地址。
    """
    urls = [UPDATE_API] + [m + UPDATE_API for m in UPDATE_API_MIRRORS]
    last_err = None
    for u in urls:
        req = urllib.request.Request(
            u,
            headers={
                "Accept": "application/vnd.github+json",
                "User-Agent": f"AIReplay-check/{current_app_version()}",
            },
        )
        try:
            with urllib.request.urlopen(req, timeout=UPDATE_TIMEOUT) as resp:
                return json.loads(resp.read().decode("utf-8")), None
        except urllib.error.HTTPError as e:
            last_err = e
            # 只有「直连」的 404 才算数（真没有 Release）；线路配额/限流一律换下一条再试
            if e.code == 404 and u == urls[0]:
                return None, "暂时没有可用的更新信息"
            continue
        except Exception as e:
            last_err = e
            continue
    # 全部线路都不通：只保留异常类型，不把可能含地址的原文带到界面上
    if isinstance(last_err, urllib.error.HTTPError):
        return None, f"更新服务返回异常（HTTP {last_err.code}）"
    return None, f"网络异常，暂时无法获取更新信息（{type(last_err).__name__}）"


def check_for_update():
    """查询更新服务上的最新版本信息。

    返回 (status, payload)：
      ("update", info)  有可用更新。info = {version, tag, name, url, asset_name, size, notes}
      ("latest", info)  已是最新版本
      ("error",  msg)   检测失败（网络不可达 / 尚无 Release 等），msg 为可读说明
    """
    data, err = _fetch_release_json()
    if err:
        return ("error", err)

    tag = (data.get("tag_name") or "").strip()
    assets = data.get("assets") or []

    asset = None
    for a in assets:  # 优先取 AIReplay-*.exe
        nm = (a.get("name") or "")
        if nm.lower().endswith(".exe") and nm.lower().startswith(UPDATE_ASSET_PREFIX.lower()):
            asset = a
            break
    if asset is None:  # 退而求其次：任意 .exe
        for a in assets:
            if (a.get("name") or "").lower().endswith(".exe"):
                asset = a
                break
    if asset is None:
        return ("error", "该版本暂无可下载的更新文件")

    info = {
        "version": tag.lstrip("vV"),
        "tag": tag,
        "name": data.get("name") or tag,
        "url": asset.get("browser_download_url") or "",
        "asset_name": asset.get("name") or "",
        "size": int(asset.get("size") or 0),
        "notes": data.get("body") or "",
    }
    if version_is_newer(info["version"], current_app_version()):
        return ("update", info)
    return ("latest", info)


def update_candidates(url):
    """按「直连 → 各加速线路」排出候选下载地址（地址本身不展示给用户）。"""
    urls = [url]
    for m in UPDATE_ASSET_MIRRORS:
        u = m + url
        if u not in urls:
            urls.append(u)
    return urls


def _download_once(url, tmp, expect, on_progress=None):
    """把 url 的内容下载 / 续传到 tmp，返回 tmp 当前的总字节数。

    支持断点续传：tmp 里已有内容时用 Range 从断点继续；
    线路不支持 Range（回 200 而不是 206）就丢掉重下。
    国内线路经常「连得上、下到一半断流」，续传能显著提高成功率。
    """
    offset = os.path.getsize(tmp) if os.path.exists(tmp) else 0
    if expect and offset >= expect:
        return offset

    headers = {
        "User-Agent": f"AIReplay-check/{current_app_version()}",
        "Accept-Encoding": "identity",     # 避免线路把包压一遍导致字节数对不上
    }
    if offset:
        headers["Range"] = f"bytes={offset}-"

    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=UPDATE_DL_TIMEOUT) as resp:
        if offset and getattr(resp, "status", 200) != 206:
            offset = 0                     # 线路不支持续传，从头写过
            mode = "wb"
        else:
            mode = "ab"
        total = int(resp.headers.get("Content-Length") or 0) + offset
        if not total:
            total = expect or 0
        got = offset
        with open(tmp, mode) as f:
            while True:
                chunk = resp.read(256 * 1024)
                if not chunk:
                    break
                f.write(chunk)
                got += len(chunk)
                if on_progress:
                    try:
                        on_progress(got, total)
                    except Exception:
                        pass
    return got


def _verify_asset(path, got, expect):
    """产物校验：字节数对得上，且确实是 Windows 可执行文件（PE 头 MZ）。

    线路返回错误页 / 半截包时（HTTP 200 也可能是网页），靠这两条挡住。
    """
    if got < UPDATE_MIN_BYTES:
        raise RuntimeError(f"下载内容不完整（仅 {got} 字节）")
    if expect and got != expect:
        raise RuntimeError(f"下载不完整：{got} / {expect} 字节")
    with open(path, "rb") as f:
        if f.read(2) != b"MZ":
            raise RuntimeError("下载内容不是有效的程序文件")


def download_update(info, on_progress=None, on_stage=None):
    """下载新版本 exe 到程序所在目录，返回新文件路径。

    两层容错，专门对付「下载失败」：
      1. 同一条线路内：断流就带着已有字节**续传**，最多试 UPDATE_DL_ATTEMPTS 次；
      2. 线路之间：直连优先，不通就自动依次切到加速线路（UPDATE_ASSET_MIRRORS）。

    全程只写 .part，校验通过才改名成正式文件；任何一次失败都不会
    把半截文件当成可用程序，也不会在目录里留下残包。
    """
    url = info.get("url") or ""
    name = info.get("asset_name") or ""
    if not url or not name:
        raise RuntimeError("下载地址或文件名缺失")

    dest = os.path.join(script_dir(), os.path.basename(name))
    tmp = dest + ".part"
    expect = int(info.get("size") or 0)

    lines = update_candidates(url)
    last_err = None
    for idx, u in enumerate(lines, 1):
        try:
            os.remove(tmp)                 # 换线路：丢弃上一条线路的残包
        except OSError:
            pass
        if idx > 1:
            if on_progress:
                try:
                    on_progress(0, expect)
                except Exception:
                    pass
            if on_stage:
                try:
                    on_stage(idx, len(lines))
                except Exception:
                    pass
        for att in range(1, UPDATE_DL_ATTEMPTS + 1):
            try:
                got = _download_once(u, tmp, expect, on_progress)
            except Exception as e:
                last_err = e
                log(f"[更新] 第 {idx}/{len(lines)} 条线路第 {att} 次中断：{type(e).__name__}")
                continue                   # 保留 .part，下次续传
            if expect and got < expect:
                last_err = RuntimeError(f"下载不完整：{got} / {expect} 字节")
                log(f"[更新] 第 {idx}/{len(lines)} 条线路第 {att} 次不完整：{got}/{expect}")
                continue                   # 保留 .part，下次续传
            try:
                _verify_asset(tmp, got, expect)
            except Exception as e:
                last_err = e
                log(f"[更新] 第 {idx}/{len(lines)} 条线路内容校验未通过，换线路")
                break                      # 内容不对，重试同一条线路没意义
            os.replace(tmp, dest)
            log(f"[更新] 新版本已下载完成: {dest}")
            return dest
        try:
            os.remove(tmp)                 # 本条线路用尽：删残包再换下一条
        except OSError:
            pass

    raise RuntimeError(
        f"下载未完成（{type(last_err).__name__}，已尝试 {len(lines)} 条线路）"
    ) from last_err


def launch_detached(path):
    """以独立进程启动指定程序（不随当前进程退出而被回收）"""
    import subprocess
    try:
        if hasattr(os, "startfile"):
            os.startfile(path)   # noqa: S606 —— Windows：等同双击运行
            return
    except OSError:
        pass
    subprocess.Popen([path], cwd=os.path.dirname(path) or None)


def prepare_update_cleanup(old_exe):
    """记录旧版本文件，交由新版本启动时清理（避免目录里堆积历史 exe）"""
    try:
        with open(os.path.join(script_dir(), UPDATE_MARKER), "w", encoding="utf-8") as f:
            json.dump({"remove": [old_exe]}, f)
    except Exception:
        pass


def cleanup_after_update():
    """启动时清理上次在线更新遗留的旧版本文件。

    新版本刚启动时旧 exe 可能仍被占用，因此带重试；仍失败则保留记录，下次启动再试。
    安全约束：只删程序所在目录下的文件，且绝不删当前正在运行的程序。
    """
    marker = os.path.join(script_dir(), UPDATE_MARKER)
    if not os.path.exists(marker):
        return
    try:
        with open(marker, encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        data = {}

    folder = os.path.abspath(script_dir())
    current = current_exe_path()
    pending = []

    for path in data.get("remove") or []:
        target = os.path.abspath(str(path))
        if os.path.dirname(target) != folder:      # 安全：只处理同目录文件
            continue
        if current and target == current:          # 安全：绝不删正在运行的程序
            continue
        for attempt in range(10):
            try:
                os.remove(target)
                break
            except FileNotFoundError:
                break
            except Exception:
                if attempt == 9:
                    pending.append(str(path))
                else:
                    time.sleep(1)

    if pending:
        try:
            with open(marker, "w", encoding="utf-8") as f:
                json.dump({"remove": pending}, f)
        except Exception:
            pass
        return
    try:
        os.remove(marker)
    except OSError:
        pass


# ==================================================================
#                        授权验证（License Guard）
# ==================================================================
# 错误码 → 给用户看的中文处理建议：宁可说清楚原因，也不要无声闪退
_LICENSE_HINTS = {
    "NO_KEY":               "本机没有找到授权码。请把授权码文件 license.key 放到程序同目录，或联系供应商。",
    "KEY_NOT_FOUND":        "授权码不存在或填写有误，请核对后重试。",
    "LICENSE_EXPIRED":      "授权已到期，请联系供应商续期。",
    "LICENSE_REVOKED":      "授权已被吊销，请联系供应商。",
    "LICENSE_SUSPENDED":    "授权已被暂停，请联系供应商。",
    "DEVICE_LIMIT":         "该授权码的可绑定设备数已用满。请在旧设备上解绑，或在管理后台移除一台设备后重试。",
    "DEVICE_NOT_ACTIVATED": "本机尚未激活，且当前不允许自动占用新的设备位。",
    "MACHINE_MISMATCH":     "授权绑定的是另一台电脑（本机状态文件可能被复制过），请联系供应商重新授权。",
    "CLOCK_TAMPERED":       "检测到系统时间异常（被调回过去），请校准系统时间后重试。",
    "OFFLINE_TOO_LONG":     "太久没能连上授权服务器，已超出离线容忍期。请连接网络后重试。",
    "NETWORK":              "无法连接授权服务器，且本机没有可用的离线授权。请检查网络后重试。",
    "BAD_SIGNATURE":        "授权令牌验签失败（公钥不匹配或文件被篡改），请联系供应商。",
    "BAD_TOKEN":            "授权令牌格式异常，请联网重新校验。",
    "APP_UNKNOWN":          "授权后台不认识当前 app_id，请核对 LICENSE_APP_ID 配置。",
    "RATE_LIMITED":         "请求过于频繁，请稍后重试。",
    "SDK_MISSING":          "程序缺少授权模块（license_guard.py），无法校验授权。",
    "CONFIG_ERROR":         "授权配置不完整，无法校验授权。",
}


def license_hint(code, message=""):
    """把授权错误码翻译成可执行的中文建议"""
    return _LICENSE_HINTS.get(str(code or "").upper()) or (message or f"授权校验未通过（{code}）")


def resolve_license_key():
    """按优先级决定使用哪个授权码，返回 (key, 来源说明)

      1. 环境变量 LICENSE_KEY —— 便于临时指定或自动化调用
      2. 程序同目录的 license.key —— 分发时一个客户一个文件
      3. 脚本内置的 LICENSE_KEY —— 本地自用
    """
    env = (os.environ.get("LICENSE_KEY") or "").strip()
    if env:
        return env, "环境变量 LICENSE_KEY"

    for p in (os.path.join(script_dir(), LICENSE_FILE_NAME),
              os.path.join(os.getcwd(), LICENSE_FILE_NAME)):
        try:
            if os.path.isfile(p):
                text = open(p, encoding="utf-8", errors="ignore").read().strip()
                if text:
                    return text.splitlines()[0].strip(), f"{p}"
        except Exception:
            continue

    builtin = (LICENSE_KEY or "").strip()
    if builtin:
        return builtin, "内置授权码"
    return "", "未找到"


def create_license_guard(on_denied=None):
    """构造 LicenseGuard；未启用授权 / SDK 缺失 / 配置异常时返回 None

    注意：on_denied 必须传入。SDK 只有在拿到 on_denied 时才会把失败「抛」回来，
    否则它会自己按 fail_mode 处理 —— hard 直接 os._exit，message 会弹一个
    系统对话框，两者都不受本程序控制（弹框会泄漏授权信息到界面）。
    """
    if not (LICENSE_ENABLED and _LICENSE_SDK_OK):
        return None
    key, _src = resolve_license_key()
    try:
        cfg = license_guard.GuardConfig(
            server_url=LICENSE_SERVER,
            app_id=str(LICENSE_APP_ID),
            public_key=LICENSE_PUBLIC_KEY,
            license_key=key,
            sign_alg=LICENSE_SIGN_ALG,
            request_timeout=LICENSE_TIMEOUT,
            retry_times=2,
            watchdog_interval_sec=LICENSE_WATCHDOG_SEC,
            heartbeat_cap_sec=LICENSE_HEARTBEAT_CAP_SEC,   # 0 = 不封顶，跟随后端心跳
            online_check_interval_sec=LICENSE_ONLINE_CHECK_SEC,  # 0 = 每次校验都联网
            fail_mode=LICENSE_FAIL_MODE,   # 兜底默认；成功后端策略会覆盖它
            exit_code=LICENSE_EXIT_CODE,   # 同上
            allow_dialog=False,            # 铁律：SDK 不得自己弹任何授权对话框
            on_denied=on_denied,
        )
        return license_guard.LicenseGuard(cfg)
    except Exception:
        return None


def license_fail_behavior(guard=None):
    """返回校验失败时该怎么做：(是否弹提示, 退出码)

    本程序永不弹授权提示：不管什么情况，都不能出现授权码相关的提示，
    也不能让人看出软件设置了授权。因此第一个返回值恒定是 False ——
    即便后台把「到期时的行为」设成「弹提示后退出」，界面上也不会出现任何对话框，
    该旋钮对界面彻底失效（后台那边仍可留着，不影响）。

    退出码依旧跟随后端下发的 exit_code，方便运维 / 定时任务统计退出原因。
    """
    cfg = getattr(guard, "cfg", None)
    code = getattr(cfg, "exit_code", None)
    exit_code = code if isinstance(code, int) and code > 0 else (LICENSE_EXIT_CODE or 1)
    return False, exit_code


def verify_license(on_denied=None):
    """执行一次授权校验。

    返回 (ok, payload, guard)：
      ok=True  → payload 为 LicenseInfo，guard 已通过校验（可用于启动看守）
      ok=False → payload 为 (错误码, 中文处理建议)，guard 仍会返回（若非 None），
                 因为它的 cfg 里带着后端下发的「到期时的行为」/ 退出码，
                 调用方要据此决定弹不弹提示、用哪个退出码
    """
    if not LICENSE_ENABLED:
        return True, None, None
    if not _LICENSE_SDK_OK:
        return False, ("SDK_MISSING", license_hint("SDK_MISSING")), None
    guard = create_license_guard(on_denied=on_denied)
    if guard is None:
        return False, ("CONFIG_ERROR", license_hint("CONFIG_ERROR")), None
    try:
        return True, guard.verify(allow_offline=True), guard
    except license_guard.LicenseError as e:
        return False, (e.code, license_hint(e.code, e.message)), guard
    except Exception as e:
        return False, ("UNKNOWN", f"授权校验异常: {e}"), guard


def enforce_license_cli():
    """命令行 / 降级弹窗路径的授权校验（图形界面模式另有自己的后台校验线程）。

    返回 0 表示通过；非 0 为应返回给系统的退出码（跟随后端下发的 exit_code）。
    **全程零输出**：以前会把失败原因打到控制台，但 exe 从终端启动时
    终端会看到这些字样，等于暴露「软件带授权」。现在只留退出码，
    需要看原因请设 AIREPLAY_LICENSE_TRACE 环境变量（见 _license_trace）。
    """
    if not LICENSE_ENABLED:
        return 0
    ok, payload, guard = verify_license(
        on_denied=lambda c, m: None)          # 先自行接管，避免 SDK 直接闪退
    if ok:
        if guard is not None:
            # 长期运行的命令行任务同样接受运行期复查：被吊销 / 到期立即退出
            _quit_code = license_fail_behavior(guard)[1]
            guard.cfg.on_denied = lambda c, m: _exit_now(_quit_code)
            try:
                guard.start_watchdog()
            except Exception:
                pass
        return 0
    code, msg = payload
    _license_trace(f"命令行校验未通过（{code}）：{msg}")
    return license_fail_behavior(guard)[1]


# ---------------------------------------------------------------- 登录 Token
JWT_RE = re.compile(rb"eyJ[A-Za-z0-9_\-+/=]+\.eyJ[A-Za-z0-9_\-+/=]+\.[0-9a-fA-F]{64}")
BROWSER_ROOTS = [
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "User Data"),
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Edge", "User Data"),
]


def parse_token_info(token):
    """解析 JWT 中间段，取 account / expire 等字段"""
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return json.loads(base64.urlsafe_b64decode(payload.encode("ascii")).decode("utf-8"))
    except Exception:
        return {}


def scan_leveldb_dir(leveldb_dir, tmp_root):
    """复制 LevelDB 文件后扫描，返回 [(expire, token, info), ...]"""
    found = []
    tmp_dir = os.path.join(tmp_root, "ls_" + str(abs(hash(leveldb_dir)) % 10 ** 8))
    os.makedirs(tmp_dir, exist_ok=True)
    for name in os.listdir(leveldb_dir):
        src = os.path.join(leveldb_dir, name)
        if not os.path.isfile(src) or name == "LOCK":
            continue
        dst = os.path.join(tmp_dir, name)
        try:
            shutil.copy2(src, dst)
            blob = open(dst, "rb").read()
        except Exception:
            continue
        for m in JWT_RE.finditer(blob):
            ctx = blob[max(0, m.start() - 600): m.start()]
            if b"amis.ssjj.cn" not in ctx and b"AMIS-USER-TOKEN" not in ctx:
                continue
            token = m.group(0).decode("ascii")
            info = parse_token_info(token)
            if info.get("account"):
                found.append((int(info.get("expire") or 0), token, info))
    return found


# ---------------------------------------------------------------- 本地 Token 缓存
# 缓存文件位于脚本 / exe 同目录，仅保存 token / account / expire / saved_at，
# 绝不保存密码；Token 失效或被服务端判定失效时自动删除，下次重新登录后再写入。
def token_cache_path():
    """本地 Token 缓存文件路径（脚本 / exe 同目录下的 .amis_token.json）"""
    return os.path.join(script_dir(), TOKEN_CACHE_NAME)


def _fmt_expire(expire):
    """把 expire 时间戳格式化为可读文本；无效时返回「未知」"""
    try:
        expire = int(expire or 0)
    except (TypeError, ValueError):
        return "未知"
    if not expire:
        return "未知"
    try:
        return datetime.fromtimestamp(expire).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return "未知"


def cache_is_fresh(info):
    """缓存 Token 是否仍在有效期内

    expire 缺失（无法解析出有效期）时视为可用：真失效会被服务端判定，
    届时自动删缓存并弹框重新登录，不会造成不可恢复的死循环。
    """
    try:
        expire = int((info or {}).get("expire") or 0)
    except (TypeError, ValueError):
        return False
    if not expire:
        return True
    return expire > datetime.now().timestamp()


def load_token_cache():
    """读取本地 Token 缓存；返回 (token, info)；无缓存/损坏时返回 None

    文件不存在：返回 None（按无缓存处理，不打印告警）。
    文件损坏或解析失败：删除该文件并按无缓存处理。
    """
    path = token_cache_path()
    if not os.path.isfile(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            raise ValueError("缓存内容不是 JSON 对象")
        token = str(data.get("token") or "").strip()
        if not token:
            raise ValueError("缓存中缺少 token 字段")
        info = parse_token_info(token)
        account = str(data.get("account") or "").strip()
        if account:
            info["account"] = account
        try:
            expire = int(data.get("expire") or 0)
        except (TypeError, ValueError):
            expire = 0
        if expire:
            info["expire"] = expire
        return token, info
    except Exception as e:
        log(f"[警告] 本地 Token 缓存损坏或无法解析，已按无缓存处理并删除: {e}")
        remove_token_cache("缓存损坏")
        return None


def save_token_cache(token, info):
    """写入本地 Token 缓存（仅 token / account / expire / saved_at，不含密码）"""
    if not token:
        return False
    info = info or {}
    try:
        expire = int(info.get("expire") or 0)
    except (TypeError, ValueError):
        expire = 0
    payload = {
        "token": token,
        "account": str(info.get("account") or ""),
        "expire": expire,
        "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    path = token_cache_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        log(f"[缓存] 已写入本地 Token 缓存（账号 {payload['account'] or '未知'}，"
            f"有效期至 {_fmt_expire(expire)}）: {path}")
        return True
    except Exception as e:
        log(f"[警告] 本地 Token 缓存写入失败（不影响本次抓取）: {e}")
        return False


def remove_token_cache(reason=""):
    """删除本地 Token 缓存文件（Token 失效 / 缓存损坏时调用）"""
    path = token_cache_path()
    if not os.path.isfile(path):
        return False
    try:
        os.remove(path)
        note = f"（{reason}）" if reason else ""
        log(f"[缓存] 已删除本地 Token 缓存{note}: {path}")
        return True
    except Exception as e:
        log(f"[警告] 本地 Token 缓存删除失败: {e}")
        return False


# ---------------------------------------------------------------- 本机加密凭据（Windows DPAPI）
# 登录框勾选「记住账号密码（本机加密保存）」且登录成功后，将账号密码经 DPAPI
# （CryptProtectData，ctypes 调用，不新增 pywin32 依赖）加密写入脚本 / exe 同目录下的
# .amis_credentials.bin，供 Token 缺失 / 过期 / 失效时自动登录使用。密码不以明文落盘，
# 也不写入日志 / 缓存 / 代码，仅在内存中传递。
class _DATA_BLOB(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", wintypes.LPVOID)]


# 必须显式声明 LocalFree 的 argtypes。不声明时 ctypes 会把 Python int 按
# 32 位 C int 转换，而 64 位进程里 DPAPI 返回的堆指针通常超过 2^31，
# LocalFree(指针) 会抛 OverflowError（int too long to convert）——
# 曾导致「加密成功但函数以异常收场」，save_credentials 只 open 没 write，
# 凭据文件被写成 0 字节，记住的账号密码从未真正落盘。
_LocalFree = ctypes.windll.kernel32.LocalFree
_LocalFree.argtypes = [ctypes.c_void_p]      # 按指针宽度传递，高位地址不再溢出
_LocalFree.restype = ctypes.c_void_p


def _dpapi_protect(data):
    """Windows DPAPI 加密 bytes，返回密文 bytes；失败抛 OSError"""
    buf = ctypes.create_string_buffer(data, len(data))
    blob_in = _DATA_BLOB(len(data), ctypes.cast(buf, wintypes.LPVOID))
    blob_out = _DATA_BLOB()
    if not ctypes.windll.crypt32.CryptProtectData(
            ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out)):
        raise OSError(f"CryptProtectData 失败（错误码 {ctypes.get_last_error()}）")
    try:
        raw = ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        _LocalFree(blob_out.pbData)
    return bytes(raw)


def _dpapi_unprotect(data):
    """Windows DPAPI 解密 bytes，返回明文 bytes；失败抛 OSError"""
    buf = ctypes.create_string_buffer(data, len(data))
    blob_in = _DATA_BLOB(len(data), ctypes.cast(buf, wintypes.LPVOID))
    blob_out = _DATA_BLOB()
    if not ctypes.windll.crypt32.CryptUnprotectData(
            ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out)):
        raise OSError(f"CryptUnprotectData 失败（错误码 {ctypes.get_last_error()}）")
    try:
        raw = ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        _LocalFree(blob_out.pbData)
    return bytes(raw)


def credentials_path():
    """本机加密凭据文件路径（脚本 / exe 同目录下的 .amis_credentials.bin）"""
    return os.path.join(script_dir(), CREDENTIALS_FILE_NAME)


def save_credentials(account, password, ldap=False):
    """加密保存账号密码到本机（供 Token 失效时自动登录）

    写入内容为 {account, password, ldap, saved_at}，整包经 DPAPI 加密。
    仅接受内存中已有的明文，不读取任何外部来源；失败只记日志，不影响登录流程。
    """
    account = str(account or "").strip()
    if not account or password is None:
        return False
    payload = {
        "account": account,
        "password": str(password),
        "ldap": bool(ldap),
        "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    # 先完成加密、再打开文件：一旦加密失败不会在磁盘上留下 0 字节的凭据文件
    # （0 字节文件下次启动会被当作「凭据损坏」删掉，等同于从未记住过密码）
    try:
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        enc = _dpapi_protect(raw)
    except Exception as e:
        log(f"[警告] 账号密码加密保存失败（不影响本次登录）: {e}")
        return False
    path = credentials_path()
    try:
        with open(path, "wb") as f:
            f.write(enc)
        log(f"[凭据] 账号密码已加密保存到本机（{path}）。")
        return True
    except Exception as e:
        log(f"[警告] 凭据文件写入失败（不影响本次登录）: {e}")
        return False


def load_credentials():
    """读取并解密本机凭据；返回 dict 或 None（无文件 / 损坏 / 解密失败时按无凭据处理）"""
    path = credentials_path()
    if not os.path.isfile(path):
        return None
    try:
        with open(path, "rb") as f:
            enc = f.read()
        data = json.loads(_dpapi_unprotect(enc).decode("utf-8"))
        if not isinstance(data, dict):
            raise ValueError("凭据内容不是 JSON 对象")
        account = str(data.get("account") or "").strip()
        password = str(data.get("password") or "")
        if not account or not password:
            raise ValueError("凭据缺少账号或密码")
        return {
            "account": account,
            "password": password,
            "ldap": bool(data.get("ldap")),
            "saved_at": str(data.get("saved_at") or ""),
        }
    except Exception as e:
        log(f"[警告] 本机加密凭据损坏或无法解密，已删除并视为未保存: {e}")
        remove_credentials("凭据损坏或无法解密")
        return None


def remove_credentials(reason=""):
    """删除本机加密凭据文件（用户取消勾选 / 凭据被服务端拒绝 / 凭据损坏时调用）"""
    path = credentials_path()
    if not os.path.isfile(path):
        return False
    try:
        os.remove(path)
        note = f"（{reason}）" if reason else ""
        log(f"[凭据] 已删除本机加密保存的账号密码{note}: {path}")
        return True
    except Exception as e:
        log(f"[警告] 本机加密凭据删除失败: {e}")
        return False


class CredentialsRejectedError(RuntimeError):
    """服务端明确拒绝本次登录凭据（账号 / 密码错误、账号停用等）"""


class TokenInvalidError(RuntimeError):
    """服务端判定登录态失效（HTTP 401/403 或返回鉴权 / 登录失效类错误码、消息）"""


# 出现以下消息关键字组合时，视为登录态失效
_AUTH_FAIL_PHRASES = (
    "未登录", "请先登录", "请重新登录", "重新登录", "登录已失效", "登录失效",
    "登录过期", "登录状态失效", "鉴权失败", "鉴权错误", "认证失败", "未授权",
    "无效的token", "token无效", "token 无效", "token已过期", "token 已过期",
    "令牌无效", "令牌已失效",
)


def looks_like_auth_failure(code=None, msg=""):
    """判断接口返回是否属于「Token / 登录态失效」类错误"""
    try:
        code = int(code)
    except (TypeError, ValueError):
        code = None
    if code in (401, 403):
        return True
    text = str(msg or "").strip()
    if not text:
        return False
    low = text.lower().replace(" ", "")
    if any(p.replace(" ", "") in low for p in _AUTH_FAIL_PHRASES):
        return True
    has_token_word = ("token" in low) or ("令牌" in text)
    has_auth_word = any(k in text for k in ("登录", "登陆", "失效", "过期", "鉴权", "认证", "授权"))
    return has_token_word and has_auth_word


def find_amis_token():
    """按优先级获取 amis.ssjj.cn 登录 Token

    顺序：环境变量 AMIS_TOKEN > 本地缓存 .amis_token.json（未过期）>
    浏览器 Chrome / Edge Local Storage（取有效期最长者）。
    返回 (token, info, source)；均不可用时返回 (None, {}, "")，由上层弹登录框。
    """
    env_token = os.environ.get("AMIS_TOKEN", "").strip()
    if env_token:
        info = parse_token_info(env_token)
        return env_token, info, "环境变量 AMIS_TOKEN"

    cached = load_token_cache()
    if cached:
        token, info = cached
        if cache_is_fresh(info):
            return token, info, "本地缓存"
        # 缓存已过期：删除后继续走后续获取路径
        remove_token_cache("本地缓存已过期")

    tmp_root = tempfile.mkdtemp(prefix="replay_token_")
    candidates = []
    try:
        for root in BROWSER_ROOTS:
            if not os.path.isdir(root):
                continue
            for profile in os.listdir(root):
                leveldb = os.path.join(root, profile, "Local Storage", "leveldb")
                if os.path.isdir(leveldb):
                    candidates.extend(scan_leveldb_dir(leveldb, tmp_root))
    finally:
        shutil.rmtree(tmp_root, ignore_errors=True)

    if not candidates:
        return None, {}, ""
    candidates.sort(key=lambda x: x[0], reverse=True)
    _, token, info = candidates[0]
    return token, info, "浏览器 Local Storage"


# ---------------------------------------------------------------- 登录兜底（Token 失效时手动登录）
# 登录框 UI 挂载点：GUI 启动时注册「主线程弹框」回调；未注册时回退原生 tkinter 对话框
_LOGIN_UI = None


def login_by_password(account, password, ldap=True):
    """调用后台登录接口换取新 Token（与界面完全解耦）

    成功返回 (token, info)；凭据被服务端拒绝（账号 / 密码错误等）时抛
    CredentialsRejectedError；网络 / HTTP 异常抛 RuntimeError，消息为错误说明。
    账号密码仅在本函数调用期间存在于内存中，不做任何持久化。
    """
    body = json.dumps({
        "account": account,
        "password": password,
        "ldap": bool(ldap),
    }).encode("utf-8")
    req = urllib.request.Request(
        LOGIN_API,
        data=body,
        headers={
            "Content-Type": "application/json",
            # 注意: /auth/signIn 的 X-PROJECT-ID 必须固定为 "1"（实测用 PROJECT_ID=28
            # 或不带该头均返回 HTTP 400 Bad Request，用 "1" 才返回正常 JSON）。
            # 数据抓取接口的 X-PROJECT-ID 仍为 PROJECT_ID(28)，不要混用。
            "X-PROJECT-ID": LOGIN_PROJECT_ID,
            "Accept": "application/json, text/plain, */*",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36",
            "Referer": LOGIN_PAGE,
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=HTTP_TIMEOUT) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"登录接口返回 HTTP {e.code}（{e.reason}）")
    except Exception as e:
        raise RuntimeError(f"登录请求发送失败：{e}")

    if payload.get("code") != 0:
        raise CredentialsRejectedError(
            payload.get("msg") or f"登录失败（code={payload.get('code')}）")

    token = str(((payload.get("data") or {}).get("token")) or "").strip()
    if not token:
        raise RuntimeError("登录成功但响应中未包含 Token（data.token 为空）")

    info = parse_token_info(token)
    if not info.get("account"):
        info["account"] = account
    return token, info


def try_auto_login():
    """用本机加密保存的凭据自动登录换取新 Token（全程不弹框）

    返回 (token, info)：自动登录成功（并已写入 Token 缓存）；
    返回 None：未保存凭据，或凭据被服务端拒绝（此时已自动删除本地凭据，
    由调用方回退到手动登录框），或网络 / 服务异常（保留本地凭据）。
    """
    creds = load_credentials()
    if not creds:
        return None
    try:
        token, info = login_by_password(creds["account"], creds["password"], creds["ldap"])
    except CredentialsRejectedError as e:
        remove_credentials("已保存凭据被服务端拒绝（账号或密码可能已变更）")
        log(f"[自动登录] 已保存凭据被服务端拒绝，已删除本机凭据，将回退手动登录: {e}")
        return None
    except Exception as e:
        log(f"[自动登录] 网络或服务异常，保留本机凭据，本次回退手动登录: {e}")
        return None
    save_token_cache(token, info)
    log("[自动登录] 使用已保存的账号密码登录成功，已刷新 Token 缓存，全程未弹框。")
    return token, info


def ask_credentials(tip="", error="", default_account=""):
    """弹出登录框收集凭据；返回 (account, password, ldap, remember)，用户取消返回 None

    remember 表示「记住账号密码（本机加密保存）」是否勾选（默认勾选）。
    UI 与登录逻辑解耦：GUI 模式下走 _LOGIN_UI 注册的深色对话框，
    其余场景回退到原生 tkinter 对话框。
    """
    ui = _LOGIN_UI
    if ui is not None:
        return ui(tip, error, default_account)
    return _ask_credentials_tk(tip, error, default_account)


def _ask_credentials_tk(tip="", error="", default_account=""):
    """无 customtkinter 时的原生 tkinter 登录对话框（兜底）"""
    import tkinter as tk

    result = {"value": None}
    root = tk.Tk()
    root.title("登录 amis.ssjj.cn")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    tk.Label(root, text="登录 amis.ssjj.cn",
             font=("Microsoft YaHei UI", 13, "bold")).grid(
        row=0, column=0, columnspan=2, sticky="w", padx=14, pady=(14, 2))
    tk.Label(root, text=tip or "自动获取 Token 失败，请手动登录后继续抓取。",
             fg="#666666", justify="left", wraplength=360).grid(
        row=1, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 8))

    tk.Label(root, text="登录网址").grid(row=2, column=0, sticky="e", padx=14, pady=6)
    url_row = tk.Frame(root)
    url_row.grid(row=2, column=1, sticky="w", padx=14, pady=6)
    tk.Label(url_row, text=LOGIN_PAGE, fg="#2F6BFF").pack(side="left")
    tk.Button(url_row, text="打开",
              command=lambda: webbrowser.open(LOGIN_PAGE)).pack(side="left", padx=(8, 0))

    tk.Label(root, text="账号").grid(row=3, column=0, sticky="e", padx=14, pady=6)
    var_account = tk.StringVar(value=default_account)
    ent_account = tk.Entry(root, textvariable=var_account, width=30)
    ent_account.grid(row=3, column=1, sticky="w", padx=14, pady=6)

    tk.Label(root, text="密码").grid(row=4, column=0, sticky="e", padx=14, pady=6)
    var_password = tk.StringVar()
    ent_password = tk.Entry(root, textvariable=var_password, width=30, show="●")
    ent_password.grid(row=4, column=1, sticky="w", padx=14, pady=6)

    var_ldap = tk.BooleanVar(value=False)
    tk.Checkbutton(root, text="LDAP 登录（本后台使用 LDAP 账号时请勾选）", variable=var_ldap).grid(
        row=5, column=0, columnspan=2, sticky="w", padx=14, pady=(2, 2))

    # 记住账号密码（本机加密保存，默认勾选；登录成功且勾选才写入 DPAPI 加密凭据）
    var_remember = tk.BooleanVar(value=True)
    tk.Checkbutton(root, text="记住账号密码（本机加密保存）", variable=var_remember).grid(
        row=6, column=0, columnspan=2, sticky="w", padx=14, pady=(2, 2))

    lbl_msg = tk.Label(root, text=error, fg="#EF4444", justify="left", wraplength=360)
    lbl_msg.grid(row=7, column=0, columnspan=2, sticky="w", padx=14, pady=(2, 4))

    def _submit():
        account = var_account.get().strip()
        password = var_password.get()
        if not account or not password:
            lbl_msg.configure(text="请输入账号和密码。")
            return
        # 凭据仅放入内存结果，不做任何持久化；是否记住由登录成功后统一处理
        result["value"] = (account, password, bool(var_ldap.get()), bool(var_remember.get()))
        root.destroy()

    def _cancel():
        result["value"] = None
        root.destroy()

    btns = tk.Frame(root)
    btns.grid(row=8, column=0, columnspan=2, sticky="e", padx=14, pady=(6, 14))
    tk.Button(btns, text="登录", width=12, command=_submit).pack(side="left", padx=(0, 8))
    tk.Button(btns, text="取消", width=8, command=_cancel).pack(side="left")

    root.protocol("WM_DELETE_WINDOW", _cancel)
    root.bind("<Return>", lambda _e: _submit())
    root.bind("<Escape>", lambda _e: _cancel())
    ent_account.focus_set()

    root.update_idletasks()
    x = (root.winfo_screenwidth() - root.winfo_width()) // 2
    y = (root.winfo_screenheight() - root.winfo_height()) // 3
    root.geometry(f"+{max(0, x)}+{max(0, y)}")
    root.mainloop()
    return result["value"]


def obtain_token_by_login(reason=""):
    """自动获取 Token 失败或已过期时，弹框手动登录换取 Token（最多重试 3 次）

    成功返回 (token, info)；用户取消或重试耗尽则抛 RuntimeError。
    账号密码仅内存使用，不写入磁盘 / 日志，脚本内也不含任何硬编码凭据。
    """
    default_account = ""
    error = ""
    for attempt in range(1, LOGIN_MAX_RETRY + 1):
        tip = reason if attempt == 1 else ""
        creds = ask_credentials(tip=tip, error=error, default_account=default_account)
        if creds is None:
            raise RuntimeError(
                "已取消登录：未取得 amis.ssjj.cn 的有效 Token，本次抓取中止。\n"
                "可先在 Chrome / Edge 登录 amis.ssjj.cn，或设置环境变量 AMIS_TOKEN 后再试。"
            )
        account, password, ldap, remember = creds
        default_account = account
        try:
            token, info = login_by_password(account, password, ldap)
        except Exception as e:
            error = str(e)
            log(f"[登录失败] 第 {attempt}/{LOGIN_MAX_RETRY} 次：{error}")
            if attempt < LOGIN_MAX_RETRY:
                continue
            raise RuntimeError(
                f"连续 {LOGIN_MAX_RETRY} 次登录失败，本次抓取中止。\n"
                f"最后一次错误：{error}\n"
                "请确认账号 / 密码 / LDAP 选项是否正确，或优先在浏览器登录 amis.ssjj.cn。"
            )
        log(f"[登录成功] 已获取 amis.ssjj.cn 登录 Token"
            f"（第 {attempt} 次尝试，LDAP={'是' if ldap else '否'}）。")
        # 记住账号密码（本机加密保存）：勾选则加密保存，取消勾选则删除已存凭据
        if remember:
            save_credentials(account, password, ldap)
        else:
            remove_credentials("登录框未勾选「记住账号密码」")
        # 登录成功统一写入本地缓存（仅 Token 与账号，密码不落盘）
        save_token_cache(token, info)
        return token, info


# ---------------------------------------------------------------- 数据抓取
def _fetch_pages(headers, start_at, end_at, progress=None):
    """按分页抓取全部记录；Token 失效时抛 TokenInvalidError，由上层换 Token 后重试"""
    items, page, total = [], 1, None
    while True:
        qs = urllib.parse.urlencode({
            "page": page,
            "perPage": PER_PAGE,
            "date_type": 2,
            "start_at": start_at,
            "end_at": end_at,
            "status": STATUS_DONE,
        })
        url = f"{HOST}{API_PATH}?{qs}"
        req = urllib.request.Request(url, headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=HTTP_TIMEOUT) as resp:
                payload = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code in (401, 403):
                raise TokenInvalidError(
                    f"接口返回 HTTP {e.code}（{e.reason}），登录态已失效"
                )
            raise RuntimeError(f"接口请求失败：HTTP {e.code}（{e.reason}）")
        except urllib.error.URLError as e:
            raise RuntimeError(f"接口请求失败：{e.reason}")
        if payload.get("code") != 0:
            msg = payload.get("msg")
            if looks_like_auth_failure(payload.get("code"), msg):
                raise TokenInvalidError(
                    f"接口返回鉴权失败（code={payload.get('code')}，msg={msg}），登录态已失效"
                )
            raise RuntimeError(f"接口返回异常: {msg or payload}")
        data = payload.get("data") or {}
        batch = data.get("items") or []
        total = data.get("total") if total is None else total
        items.extend(batch)
        log(f"[抓取] 第 {page} 页 {len(batch)} 条，累计 {len(items)}/{total} 条")
        if progress:
            progress(len(items), int(total or 0))
        if not batch or (total is not None and len(items) >= int(total)):
            break
        page += 1
    return items


def _log_token_source(source, info):
    """日志提示本次 Token 的来源与有效期"""
    expire = int((info or {}).get("expire") or 0)
    detail = f"，有效期至 {_fmt_expire(expire)}" if expire else ""
    log(f"[登录态] 已自动获取有效 Token（来源：{source or '未知'}{detail}）。")


def fetch_day(day, progress=None):
    """抓取指定日期 00:00:00 ~ 次日 00:00:00 的已核查记录

    Token 获取顺序：环境变量 AMIS_TOKEN > 本地缓存 .amis_token.json（未过期）>
    浏览器 Local Storage 有效 Token > 本机加密凭据自动登录 > 弹出登录框。
    Token 缺失 / 已过期 / 被服务端判定失效时，优先用已保存的账号密码自动登录换
    新 Token（全程不弹框）；自动登录不可用（未保存凭据 / 凭据被拒 / 网络异常）时
    才弹出登录框手动登录，登录成功后覆盖写入缓存并继续本次抓取，无需用户重跑。
    """
    token, info, source = find_amis_token()
    reason = ""
    if not token:
        reason = ("未能自动获取 amis.ssjj.cn 登录 Token"
                  "（未设置环境变量 AMIS_TOKEN，本地缓存与浏览器中也没有找到可用 Token）")
    else:
        expire = int(info.get("expire") or 0)
        if expire and expire < datetime.now().timestamp():
            reason = "自动获取到的 amis.ssjj.cn 登录 Token 已过期"

    if reason:
        # 优先用已保存凭据自动登录换新 Token，不成功才弹框
        auto = try_auto_login()
        if auto:
            token, info = auto
            source = "自动登录（已保存凭据）"
        else:
            log(f"[提示] {reason}，自动登录未成功，请在登录窗口中输入账号密码换取新 Token。")
            token, info = obtain_token_by_login(reason)
            source = "登录框"
    else:
        _log_token_source(source, info)

    account = info.get("account") or "xhdong"

    start_at = f"{day[0:4]}-{day[4:6]}-{day[6:8]} 00:00:00"
    end_at = (datetime.strptime(day, "%Y%m%d") + timedelta(days=1)).strftime("%Y-%m-%d 00:00:00")

    items = []
    auth_retry = 0
    while True:
        headers = {
            "X-AMIS-TOKEN": token,
            "X-AMIS-ACCOUNT": account,
            "X-PROJECT-ID": PROJECT_ID,
            "Accept": "application/json, text/plain, */*",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36",
            "Referer": HOST + "/",
        }
        try:
            items = _fetch_pages(headers, start_at, end_at, progress)
            break
        except TokenInvalidError as e:
            log(f"[登录态失效] {e}")
            # 服务端判定失效：删除本地缓存，先自动登录换新 Token，不成功才弹框
            remove_token_cache("服务端判定登录态失效")
            if auth_retry >= LOGIN_MAX_RETRY:
                raise RuntimeError(
                    f"登录 Token 连续 {LOGIN_MAX_RETRY} 次被服务端判定失效，本次抓取中止。\n"
                    f"最后一次错误：{e}"
                )
            auth_retry += 1
            auto = try_auto_login()
            if auto:
                token, info = auto
                account = info.get("account") or account
            else:
                log("[提示] 缓存 / 自动获取的 Token 已失效，自动登录未成功，"
                    "请在登录窗口中重新登录以换取新 Token。")
                token, info = obtain_token_by_login("自动获取的登录 Token 已失效，请重新登录")
                account = info.get("account") or account
            # 登录成功已覆盖写入缓存，继续本次抓取
    return items, start_at, end_at


def item_to_row(it):
    return [
        it.get("id", ""),
        STATUS_TEXT.get(it.get("status"), str(it.get("status", ""))),
        it.get("role_name", ""),
        it.get("role_id", ""),
        it.get("wd_id", ""),
        it.get("room_id", ""),
        it.get("score", ""),
        it.get("start_date", ""),
        it.get("end_date", ""),
        it.get("duration", ""),
        it.get("report_at", ""),
        it.get("operator", ""),
        it.get("deal_at", ""),
        "已封禁" if it.get("forbid") else "未封禁",
        it.get("remarks", ""),
    ]


def write_csv(rows, out_path):
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_ALL)
        writer.writerow(CSV_HEADERS)
        writer.writerows(rows)
    return out_path


# ---------------------------------------------------------------- 文档处理（逻辑内联自 replay核查.py）
def _column_error(col_keyword, step):
    msg = f"错误：【{step}】未找到包含「{col_keyword}」的列，请检查表头名称是否正确。"
    log(msg)
    raise RuntimeError(msg)


def detect_csv_encoding(path, sample=8192):
    for enc in ("utf-8-sig", "gb18030", "utf-8", "gbk"):
        try:
            with open(path, "r", encoding=enc) as f:
                f.read(sample)
            return enc
        except (UnicodeDecodeError, UnicodeError):
            continue
    return "latin-1"


def find_col_by_keyword(headers, keyword):
    kw = keyword.lower()
    for h in headers:
        if kw in str(h).strip().lower():
            return str(h).strip()
    return None


def resolve_text_columns(headers, ncols):
    from openpyxl.utils import column_index_from_string
    idx_set = set()
    for i, h in enumerate(headers):
        h_str = str(h).strip().lower()
        if any(kw.lower() in h_str for kw in TEXT_COLUMN_KEYWORDS):
            idx_set.add(i)
    for letter in TEXT_COLUMN_LETTERS:
        ci = column_index_from_string(letter) - 1
        if 0 <= ci < ncols:
            idx_set.add(ci)
    return idx_set


def resolve_datetime_columns(headers):
    idx_set = set()
    for i, h in enumerate(headers):
        h_str = str(h).strip().lower()
        if any(kw.lower() in h_str for kw in DATETIME_COLUMN_KEYWORDS):
            idx_set.add(i)
    return idx_set


def auto_fit_column_width(ws, padding=4, min_width=8, max_width=50):
    from openpyxl.utils import get_column_letter
    for col_cells in ws.iter_cols():
        col_letter = get_column_letter(col_cells[0].column)
        max_len = 0
        for cell in col_cells:
            val = cell.value
            text = "" if val is None else str(val)
            fullwidth = len(re.findall(r"[\u4e00-\u9fff\uff00-\uffef]", text))
            ascii_len = len(text) - fullwidth
            max_len = max(max_len, fullwidth * 2 + ascii_len)
        ws.column_dimensions[col_letter].width = min(max_width, max(min_width, max_len + padding))


# ---------------------------------------------------------------- 剪贴板（输出后自动复制指定区域）
# 只用 ctypes 调 Windows 剪贴板 API，不新增 pywin32 / pyperclip 依赖。
# 之所以同时写「文本 + HTML」两种格式：Excel 系粘贴走文本（TSV 按单元格铺开），
# 微信 / 飞书 / Word 粘贴走 HTML（直接出表格），一份数据两处都能用。
CLIPBOARD_CF_UNICODETEXT = 13
CLIPBOARD_HTML_CSS = (
    "border-collapse:collapse;font-family:'Microsoft YaHei',sans-serif;font-size:10pt;"
)
CLIPBOARD_CELL_CSS = "padding:2px 6px 2px 6px;border:1px solid #B7BEC9;"


def _clip_cell_text(value):
    """把单元格值转成粘贴出去应当看到的文字（与 Excel 里显示的一致）"""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float):
        # 分数等数值列：整数别粘成 "12.0"；非整数用 15 位有效数字，避免 %g 把大数截成科学计数法
        return str(int(value)) if value.is_integer() else f"{value:.15g}"
    return str(value)


def read_xlsx_range(xlsx_path, cell_range):
    """读取 xlsx 指定区域的文字矩阵（已去掉尾部空行 / 空列）"""
    from openpyxl import load_workbook
    from openpyxl.utils import range_boundaries

    min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    wb = load_workbook(xlsx_path, data_only=True)
    try:
        ws = wb.active
        rows = [
            [_clip_cell_text(ws.cell(row=r, column=c).value)
             for c in range(min_col, max_col + 1)]
            for r in range(min_row, max_row + 1)
        ]
    finally:
        wb.close()

    # 区域是写死的 C1:L11，当天数据不足时会剩一片空单元格：尾部空行 / 空列一律裁掉，
    # 否则粘到聊天窗口里会带出一串空白制表符。
    while rows and not any(cell.strip() for cell in rows[-1]):
        rows.pop()
    ncols = max((max((i + 1 for i, cell in enumerate(row) if cell.strip()), default=0)
                 for row in rows), default=0)
    return [row[:ncols] for row in rows]


def _cf_html_document(rows):
    """构造 Windows「HTML Format」剪贴板数据（带正确的字节偏移头）"""
    def esc(text):
        return (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

    body = "".join(
        "<tr>" + "".join(f'<td style="{CLIPBOARD_CELL_CSS}">{esc(cell)}</td>'
                         for cell in row) + "</tr>"
        for row in rows
    )
    fragment = (
        f'<table border="1" cellspacing="0" cellpadding="0" '
        f'style="{CLIPBOARD_HTML_CSS}border:1px solid #B7BEC9;">{body}</table>'
    )
    header_tpl = (
        "Version:1.0\r\n"
        "StartHTML:{sh}\r\n"
        "EndHTML:{eh}\r\n"
        "StartFragment:{sf}\r\n"
        "EndFragment:{ef}\r\n"
    )
    prefix = "<html><body>\r\n<!--StartFragment-->"
    suffix = "<!--EndFragment-->\r\n</body></html>"

    # 偏移量按「字节」算，且必须用 10 位十进制占位符撑住宽度，否则数字变长会把自己挤歪
    zeros = "0" * 10
    start_html = len(header_tpl.format(sh=zeros, eh=zeros, sf=zeros, ef=zeros).encode("utf-8"))
    start_frag = start_html + len(prefix.encode("utf-8"))
    end_frag = start_frag + len(fragment.encode("utf-8"))
    end_html = end_frag + len(suffix.encode("utf-8"))
    header = header_tpl.format(
        sh=f"{start_html:010d}", eh=f"{end_html:010d}",
        sf=f"{start_frag:010d}", ef=f"{end_frag:010d}",
    )
    return (header + prefix + fragment + suffix).encode("utf-8")


def _win_clipboard_write(items):
    """把 [(格式ID, bytes)] 一次性写入 Windows 剪贴板（多格式互不覆盖）"""
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    # 显式声明参数 / 返回类型：64 位下句柄超过 2^31，不声明会被截断成 32 位整数
    # （同 DPAPI 那次 LocalFree 的坑），导致 SetClipboardData 拿到错误的句柄。
    user32.OpenClipboard.argtypes = [wintypes.HWND]
    user32.OpenClipboard.restype = wintypes.BOOL
    user32.EmptyClipboard.argtypes = []
    user32.EmptyClipboard.restype = wintypes.BOOL
    user32.CloseClipboard.argtypes = []
    user32.CloseClipboard.restype = wintypes.BOOL
    user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
    user32.SetClipboardData.restype = wintypes.HANDLE
    kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
    kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalUnlock.restype = wintypes.BOOL
    kernel32.GlobalFree.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalFree.restype = wintypes.HGLOBAL

    # 剪贴板是全局独占资源：别的程序（Excel、输入法、截图工具）短暂占用很常见，重试几次
    opened = False
    for _ in range(10):
        if user32.OpenClipboard(None):
            opened = True
            break
        time.sleep(0.05)
    if not opened:
        raise OSError("剪贴板正被其他程序占用")

    GMEM_MOVEABLE = 0x0002
    try:
        if not user32.EmptyClipboard():
            raise OSError("清空剪贴板失败")
        for fmt, data in items:
            handle = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data))
            if not handle:
                raise OSError("剪贴板内存分配失败")
            ptr = kernel32.GlobalLock(handle)
            if not ptr:
                kernel32.GlobalFree(handle)
                raise OSError("剪贴板内存锁定失败")
            ctypes.memmove(ptr, data, len(data))
            kernel32.GlobalUnlock(handle)
            # 交给系统后所有权转移，不能再 GlobalFree
            if not user32.SetClipboardData(fmt, handle):
                kernel32.GlobalFree(handle)
                raise OSError(f"写入剪贴板格式 {fmt} 失败")
    finally:
        user32.CloseClipboard()


def set_clipboard(rows):
    """把文字矩阵写入剪贴板（文本 + HTML 表格）；失败抛 OSError"""
    tsv = "\r\n".join("\t".join(row) for row in rows)
    if os.name != "nt":
        import tkinter as tk
        root = tk.Tk()
        try:
            root.withdraw()
            root.clipboard_clear()
            root.clipboard_append(tsv)
            root.update()
        finally:
            root.destroy()
        return

    import ctypes
    items = []
    fmt_html = ctypes.windll.user32.RegisterClipboardFormatW("HTML Format")
    if fmt_html:
        items.append((fmt_html, _cf_html_document(rows)))
    # 文本格式放最后写入（高保真格式优先），UTF-16LE + 结尾双字节 0
    items.append((CLIPBOARD_CF_UNICODETEXT, tsv.encode("utf-16-le") + b"\x00\x00"))
    _win_clipboard_write(items)


def copy_xlsx_range_to_clipboard(xlsx_path, cell_range=None):
    """把 xlsx 指定区域复制到剪贴板，返回状态字典（**不抛异常**）

    剪贴板被占用 / 文件被 Excel 独占打开等都不影响已经生成的报表，
    只记一条警告，由调用方决定是否提示。
    """
    cfg = load_settings()
    if cell_range is None:
        cell_range = cfg.get("clipboard_range") or CLIPBOARD_RANGE
    info = {"ok": False, "range": cell_range, "rows": 0, "cols": 0, "error": ""}
    if not cfg.get("clipboard_enabled", CLIPBOARD_ENABLED):
        info["error"] = "已在设置中关闭"
        return info
    if not cell_range:
        info["error"] = "未配置复制区域"
        return info
    try:
        rows = read_xlsx_range(xlsx_path, cell_range)
        if not rows or not rows[0]:
            info["error"] = f"区域 {cell_range} 没有内容"
            return info
        set_clipboard(rows)
    except Exception as e:
        info["error"] = f"{type(e).__name__}: {e}"
        return info
    info.update(ok=True, rows=len(rows), cols=len(rows[0]))
    return info


def process_csv(path, out_folder, day):
    """CSV → 筛选/排序/去重 → xlsx，返回统计信息与输出路径"""
    import pandas as pd

    cfg = load_settings()
    enc = detect_csv_encoding(path)
    header_df = pd.read_csv(path, nrows=0, encoding=enc)
    headers = list(header_df.columns)
    text_cols = resolve_text_columns(headers, len(headers))
    datetime_cols = resolve_datetime_columns(headers)
    if not text_cols:
        _column_error("'id' 或 房间", "文本列")

    converters = {headers[i]: str for i in text_cols}
    for i in datetime_cols:
        converters[headers[i]] = pd.to_datetime

    df = pd.read_csv(path, encoding=enc, converters=converters, keep_default_na=False)

    # 筛选（条件来自设置页，默认「未封禁 + 无异常」）
    rows_before = len(df)
    filters = cfg.get("filters") or dict(FILTER_COLUMNS)
    for col_key, val_key in filters.items():
        col = find_col_by_keyword(df.columns, col_key)
        if col is None:
            _column_error(col_key, "筛选")
        df = df[df[col].astype(str).str.strip() == str(val_key).strip()]
    rows_after_filter = len(df)

    # 排序
    sort_col = find_col_by_keyword(df.columns, cfg.get("sort_keyword") or SORT_COLUMN_KEYWORD)
    if sort_col is None:
        _column_error(cfg.get("sort_keyword") or SORT_COLUMN_KEYWORD, "排序")
    df[sort_col] = pd.to_numeric(df[sort_col], errors="coerce")
    df = df.sort_values(by=sort_col, ascending=not cfg.get("sort_desc", True), na_position="last")

    # 去重
    dedup_key = cfg.get("dedup_keyword") or DEDUP_COLUMN_KEYWORD
    dedup_col = find_col_by_keyword(df.columns, dedup_key)
    rows_before_dedup = len(df)
    if dedup_col is None:
        _column_error(dedup_key, "去重")
    df = df.drop_duplicates(subset=[dedup_col], keep="first")
    rows_after_dedup = len(df)

    df = df.reset_index(drop=True)

    # 输出路径
    os.makedirs(out_folder, exist_ok=True)
    out_path = os.path.join(out_folder, XLSX_NAME_TPL.format(day=day))
    backup_if_exists(out_path)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        ws = writer.sheets["Sheet1"]
        for ci in text_cols:
            for r in range(2, ws.max_row + 1):
                cell = ws.cell(row=r, column=ci + 1)
                if cell.value is None or (
                    isinstance(cell.value, str) and cell.value.strip().lower() == "nan"
                ):
                    cell.value = ""
                else:
                    cell.value = str(cell.value)
                cell.number_format = "@"
        for ci in datetime_cols:
            for r in range(2, ws.max_row + 1):
                cell = ws.cell(row=r, column=ci + 1)
                if cell.value is None:
                    cell.value = ""
                cell.number_format = "yyyy-mm-dd hh:mm:ss"
        from openpyxl.styles import Alignment
        center = Alignment(horizontal="center", vertical="center")
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.alignment = center
        auto_fit_column_width(ws)

    # 输出后直接把指定区域复制到剪贴板：不打开文件、不弹 Excel，复制完即可粘贴。
    # 区域固定为 CLIPBOARD_RANGE（默认 C1:L11），数据不足时自动裁掉尾部空行空列。
    clip = copy_xlsx_range_to_clipboard(out_path)
    if clip["ok"]:
        log(f"[剪贴板] 已复制 {clip['range']}（{clip['rows']} 行 × {clip['cols']} 列）"
            f"到剪贴板，可直接粘贴到 Excel / 微信 / 文档。")
    else:
        log(f"[警告] 复制 {clip['range']} 到剪贴板未完成：{clip['error']}"
            f"（报表已正常生成，可手动打开后复制）")

    return {
        "encoding": enc,
        "rows_before": rows_before,
        "rows_after_filter": rows_after_filter,
        "rows_before_dedup": rows_before_dedup,
        "rows_after_dedup": rows_after_dedup,
        "out_path": out_path,
        "clip": clip,
    }


# ---------------------------------------------------------------- 主流程（界面与命令行共用）
class PipelineError(RuntimeError):
    """带阶段标识的流程异常（用于 Sentry 打标签与界面提示）"""

    def __init__(self, stage, message):
        super().__init__(message)
        self.stage = stage


def run_pipeline(day, progress=None):
    """抓取 → CSV → 处理 → xlsx，返回结果字典（异常转 PipelineError）

    progress: 可选回调 progress(stage, done, total)，用于界面上的阶段进度：
        ("fetch", 已抓条数, 总条数) / ("csv", 0, 0) / ("process", 0, 0) / ("clip", 0, 0)
    """
    base = script_dir()
    csv_folder = os.path.join(base, "csv")
    xlsx_folder = os.path.join(base, "replay核查")

    def _step(stage, done=0, total=0):
        if progress:
            try:
                progress(stage, done, total)
            except Exception:
                pass

    log(f"===== 目标日期: {day} =====")

    # 1. 抓取
    _step("fetch", 0, 0)
    try:
        items, start_at, end_at = fetch_day(
            day, progress=lambda done, total: _step("fetch", done, total))
    except Exception as e:
        report_exception(e, stage="fetch", day=str(day))
        log(f"[错误] 抓取失败: {e}")
        raise PipelineError("fetch", str(e))
    if not items:
        log(f"[结束] {start_at} ~ {end_at} 没有「已核查」记录。")
        return {"status": "empty", "day": day, "start_at": start_at, "end_at": end_at}

    # 2. 导出 CSV
    _step("csv")
    try:
        csv_path = write_csv(
            [item_to_row(it) for it in items],
            os.path.join(csv_folder, CSV_NAME_TPL.format(day=day)),
        )
    except Exception as e:
        report_exception(e, stage="csv", day=str(day))
        log(f"[错误] CSV 导出失败: {e}")
        raise PipelineError("csv", str(e))
    log(f"[抓取完成] 共 {len(items)} 条 → {csv_path}")

    # 3. 处理文档
    _step("process")
    try:
        stat = process_csv(csv_path, xlsx_folder, day)
    except Exception as e:
        report_exception(e, stage="process", day=str(day))
        log(f"[错误] 处理失败: {e}")
        raise PipelineError("process", str(e))

    _step("clip")
    log(f"[文档处理完成] {stat['out_path']}")
    return {
        "status": "ok",
        "day": day,
        "count": len(items),
        "start_at": start_at,
        "end_at": end_at,
        "csv_path": csv_path,
        "stat": stat,
    }


# ==================================================================
#                    用户配置（.aireplay_config.json）
# ==================================================================
# 筛选 / 排序 / 去重 / 剪贴板这些「业务参数」以前写死在代码里，改一次要改代码 + 发版。
# 现在在设置页可改，落到 exe 同目录的 .aireplay_config.json：
#   · 文件不存在 → 全部用下面 DEFAULT_SETTINGS（与老版本行为完全一致，老用户零感知）
#   · 文件损坏 / 字段缺失 → 缺什么补什么，不覆盖整个配置
#   · 已加进 .gitignore，不进仓库
CONFIG_NAME = ".aireplay_config.json"

DEFAULT_SETTINGS = {
    "filters": dict(FILTER_COLUMNS),          # {列名关键字: 期望值}
    "sort_keyword": SORT_COLUMN_KEYWORD,
    "sort_desc": True,
    "dedup_keyword": DEDUP_COLUMN_KEYWORD,
    "clipboard_enabled": CLIPBOARD_ENABLED,
    "clipboard_range": CLIPBOARD_RANGE,
    "popup_on_done": False,                   # 完成后弹结果窗（默认关：不打断粘贴）
    "animations": True,                       # 界面动效（进度条平滑 / 数字滚动 / 状态呼吸）
    "theme": "light",
    "history_limit": 80,
}

_SETTINGS_CACHE = None


def settings_path():
    return os.path.join(script_dir(), CONFIG_NAME)


def load_settings(force=False):
    """读取配置并与默认值合并；缓存于内存，避免每次处理都读盘"""
    global _SETTINGS_CACHE
    if _SETTINGS_CACHE is not None and not force:
        return _SETTINGS_CACHE
    cfg = json.loads(json.dumps(DEFAULT_SETTINGS))   # 深拷贝默认值
    path = settings_path()
    try:
        if os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                if isinstance(data.get("filters"), dict) and data["filters"]:
                    cfg["filters"] = {str(k): str(v) for k, v in data["filters"].items()}
                for key in ("sort_keyword", "dedup_keyword", "clipboard_range", "theme"):
                    if isinstance(data.get(key), str) and data[key].strip():
                        cfg[key] = data[key].strip()
                for key in ("sort_desc", "clipboard_enabled", "popup_on_done", "animations"):
                    if isinstance(data.get(key), bool):
                        cfg[key] = data[key]
                if isinstance(data.get("history_limit"), int) and data["history_limit"] > 0:
                    cfg["history_limit"] = data["history_limit"]
    except Exception as e:
        log(f"[警告] 配置文件读取失败，本次使用默认设置: {e}")
    _SETTINGS_CACHE = cfg
    return cfg


def save_settings(patch):
    """把 patch 合并进当前配置并写盘；返回合并后的配置（写盘失败只记日志）"""
    global _SETTINGS_CACHE
    cfg = load_settings()
    cfg.update(patch or {})
    _SETTINGS_CACHE = cfg
    path = settings_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        log(f"[设置] 配置已保存: {path}")
    except Exception as e:
        log(f"[警告] 配置保存失败（本次仍按新设置运行）: {e}")
    return cfg


# ==================================================================
#                     输出产物扫描（历史页 / 最近任务）
# ==================================================================
_OUTPUT_NAME_RE = None
_ROW_COUNT_CACHE = {}


def output_folder():
    return os.path.join(script_dir(), "replay核查")


def scan_outputs(limit=80, with_rows=False):
    """扫描输出目录里的报表，按日期倒序返回 [{path, day, size, mtime, rows}]"""
    global _OUTPUT_NAME_RE
    if _OUTPUT_NAME_RE is None:
        # 兼容「重命名保留旧文件」产生的 replay核查-20261118（旧-20261119_101010）.xlsx
        _OUTPUT_NAME_RE = re.compile(r"^replay核查-(\d{8})")
    folder = output_folder()
    items = []
    try:
        names = os.listdir(folder)
    except Exception:
        return items
    for name in names:
        if not name.lower().endswith(".xlsx"):
            continue
        m = _OUTPUT_NAME_RE.match(name)
        if not m:
            continue
        path = os.path.join(folder, name)
        try:
            st = os.stat(path)
        except OSError:
            continue
        items.append({
            "path": path,
            "day": m.group(1),
            "size": st.st_size,
            "mtime": st.st_mtime,
            "rows": _ROW_COUNT_CACHE.get((path, int(st.st_mtime))),
        })
    items.sort(key=lambda x: (x["day"], x["mtime"]), reverse=True)
    return items[:limit] if limit else items


def count_sheet_rows(path):
    """报表有效行数（含表头）；读取失败返回 None。结果按 (路径, mtime) 缓存"""
    try:
        st = os.stat(path)
    except OSError:
        return None
    key = (path, int(st.st_mtime))
    if key in _ROW_COUNT_CACHE:
        return _ROW_COUNT_CACHE[key]
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb.active
            n = ws.max_row or 0
        finally:
            wb.close()
    except Exception:
        n = None
    _ROW_COUNT_CACHE[key] = n
    return n


# ==================================================================
#                          图形界面（customtkinter）
# ==================================================================
try:
    import customtkinter as ctk
except Exception:  # 未安装 customtkinter 时自动退回原生弹窗模式
    ctk = None

# 配色（浅色, 深色）
CLR = {
    "win":      ("#F3F5FA", "#0F1216"),
    "card":     ("#FFFFFF", "#181B21"),
    "line":     ("#E4E8F0", "#242830"),
    "text":     ("#111827", "#EEF1F6"),
    "sub":      ("#667085", "#8B94A6"),
    "accent":   ("#2F6BFF", "#3B82F6"),
    "accent_h": ("#1E54E0", "#2C6FE0"),
    "ghost":    ("#ECF0F7", "#222731"),
    "ghost_h":  ("#DFE6F2", "#2B313D"),
    "log_bg":   ("#FAFBFD", "#0B0D11"),
    # ---- 双栏工作台新增 ----
    "side":      ("#F7F8FC", "#0D1014"),   # 左侧导航底色
    "nav_hover": ("#EDF1F9", "#1A1E26"),
    "nav_on":    ("#E7EEFF", "#1B2537"),   # 当前页
    "tile":      ("#F7F8FC", "#151920"),   # 统计块
    "stripe":    ("#F6F8FC", "#141821"),   # 预览表隔行
    "ok_bg":     ("#ECFDF3", "#0E2A1E"),   # 已复制条
    "ok_line":   ("#A7F3D0", "#14532D"),
    "ok_text":   ("#15803D", "#6EE7B7"),
}
OK_CLR, WARN_CLR, ERR_CLR, RUN_CLR = "#22C55E", "#F59E0B", "#EF4444", "#3B82F6"

if ctk is not None:

    class ResultWindow(ctk.CTkToplevel):
        """结果卡片：展示统计信息与输出文件，可一键打开目录"""

        def __init__(self, master, res):
            super().__init__(master)
            self.res = res
            ok = res.get("status") == "ok"
            stat = res.get("stat") or {}

            self.title("核查完成" if ok else "无数据")
            self.configure(fg_color=CLR["win"])
            self.resizable(False, False)
            # 高度按内容给足：7 行统计 + 输出文件框 + 底部按钮，实测内容约需 500px，
            # 卡在 470 会把「打开目录 / 再次复制 / 关闭」压成一条缝（按钮被裁掉）。
            self.geometry("640x540")

            f_title = ctk.CTkFont(family=FONT_UI, size=19, weight="bold")
            f_key = ctk.CTkFont(family=FONT_UI, size=12)
            f_val = ctk.CTkFont(family=FONT_UI, size=13, weight="bold")
            f_mono = ctk.CTkFont(family="Consolas", size=11)

            head = ctk.CTkFrame(self, fg_color="transparent")
            head.pack(fill="x", padx=24, pady=(22, 6))
            ctk.CTkLabel(
                head, text=("✓  核查完成" if ok else "!  该日期无数据"),
                font=f_title, text_color=(OK_CLR if ok else WARN_CLR),
            ).pack(side="left")

            body = ctk.CTkFrame(self, fg_color=CLR["card"], corner_radius=14,
                                border_width=1, border_color=CLR["line"])
            body.pack(fill="both", expand=True, padx=24, pady=(10, 0))
            body.grid_columnconfigure(1, weight=1)

            if ok:
                clip = stat.get("clip") or {}
                if clip.get("ok"):
                    clip_text = (f"{clip.get('range')}（{clip.get('rows')} 行 × "
                                 f"{clip.get('cols')} 列）已自动复制")
                elif clip.get("error"):
                    clip_text = f"未复制（{clip.get('error')}）"
                else:
                    clip_text = "未启用"
                rows = [
                    ("抓取日期", res.get("day", "")),
                    ("抓取范围", f"{res.get('start_at','')}  ~  {res.get('end_at','')}"),
                    ("抓取条数", f"{res.get('count', 0)} 条"),
                    ("筛选结果", f"{stat.get('rows_before', 0)}  →  {stat.get('rows_after_filter', 0)} 行"
                                 f"（未封禁 + 无异常）"),
                    ("去重结果", f"{stat.get('rows_before_dedup', 0)}  →  {stat.get('rows_after_dedup', 0)} 行"
                                 f"（按角色ID）"),
                    ("排序方式", "按分数降序"),
                    ("剪贴板", clip_text),
                ]
            else:
                rows = [
                    ("抓取日期", res.get("day", "")),
                    ("抓取范围", f"{res.get('start_at','')}  ~  {res.get('end_at','')}"),
                    ("状态", "该时间段内没有「已核查」记录"),
                ]
            for i, (k, v) in enumerate(rows):
                ctk.CTkLabel(body, text=k, font=f_key, text_color=CLR["sub"]).grid(
                    row=i, column=0, sticky="w", padx=(20, 16), pady=(14 if i == 0 else 7, 0))
                ctk.CTkLabel(body, text=v, font=f_val, text_color=CLR["text"], anchor="w",
                             justify="left", wraplength=430).grid(
                    row=i, column=1, sticky="w", padx=(0, 20), pady=(14 if i == 0 else 7, 0))

            out_path = stat.get("out_path") or res.get("csv_path") or ""
            if out_path:
                ctk.CTkLabel(body, text="输出文件", font=f_key, text_color=CLR["sub"]).grid(
                    row=len(rows), column=0, columnspan=2, sticky="w", padx=20, pady=(18, 4))
                ent = ctk.CTkEntry(body, height=34, font=f_mono, corner_radius=9,
                                   fg_color=CLR["log_bg"], border_color=CLR["line"],
                                   text_color=CLR["text"])
                ent.grid(row=len(rows) + 1, column=0, columnspan=2, sticky="ew", padx=20, pady=(0, 18))
                ent.insert(0, out_path)
                ent.configure(state="readonly")

            foot = ctk.CTkFrame(self, fg_color="transparent")
            foot.pack(fill="x", padx=24, pady=18)
            if out_path:
                ctk.CTkButton(
                    foot, text="打开输出目录", width=150, height=40, corner_radius=10,
                    font=ctk.CTkFont(family=FONT_UI, size=13, weight="bold"),
                    fg_color=CLR["accent"], hover_color=CLR["accent_h"], text_color="#FFFFFF",
                    command=lambda p=os.path.dirname(out_path): self._open(p),
                ).pack(side="left")
            # 复制是「主动动作」，粘贴前想再取一次时不用重新跑一遍抓取
            self._out_path = out_path
            self._btn_clip = None
            if out_path and CLIPBOARD_ENABLED and CLIPBOARD_RANGE:
                self._btn_clip = ctk.CTkButton(
                    foot, text=f"再次复制 {CLIPBOARD_RANGE}", width=160, height=40,
                    corner_radius=10, font=ctk.CTkFont(family=FONT_UI, size=13),
                    fg_color=CLR["ghost"], hover_color=CLR["ghost_h"], text_color=CLR["text"],
                    command=self._copy_again,
                )
                self._btn_clip.pack(side="left", padx=(10, 0))
            ctk.CTkButton(
                foot, text="关闭", width=100, height=40, corner_radius=10,
                font=ctk.CTkFont(family=FONT_UI, size=13),
                fg_color=CLR["ghost"], hover_color=CLR["ghost_h"], text_color=CLR["text"],
                command=self.destroy,
            ).pack(side="right")

            self.transient(master)
            self.after(120, self._place)
            self.lift()
            try:
                self.grab_set()
            except Exception:
                pass

        def _place(self):
            try:
                self.update_idletasks()
                m = self.master
                x = m.winfo_rootx() + (m.winfo_width() - self.winfo_width()) // 2
                y = m.winfo_rooty() + (m.winfo_height() - self.winfo_height()) // 3
                self.geometry(f"+{max(0, x)}+{max(0, y)}")
            except Exception:
                pass

        def _open(self, path):
            try:
                open_folder(path)
            except Exception as e:
                log(f"[警告] 打开目录失败: {e}")

        def _copy_again(self):
            """按当前报表重新复制一次指定区域（区域内容固定，不重新抓取）"""
            clip = copy_xlsx_range_to_clipboard(self._out_path)
            if not (clip["ok"] and self._btn_clip is not None):
                log(f"[警告] 复制 {clip['range']} 到剪贴板失败：{clip['error']}")
                return
            log(f"[剪贴板] 已复制 {clip['range']}（{clip['rows']} 行 × {clip['cols']} 列）到剪贴板。")
            try:
                self._btn_clip.configure(text="已复制 ✓")
                self.after(1600, lambda: self._btn_clip.configure(
                    text=f"再次复制 {CLIPBOARD_RANGE}"))
            except Exception:
                pass

    class LoginDialog(ctk.CTkToplevel):
        """登录对话框（深色）：账号 + 密码（掩码）+ LDAP 复选（默认不勾选）
        + 记住账号密码（本机加密保存）复选（默认勾选）

        只负责收集界面输入，返回 (account, password, ldap, remember)；用户取消返回 None。
        实际的登录请求由 login_by_password() 完成，二者互不耦合。
        """

        def __init__(self, master, tip="", error="", default_account=""):
            super().__init__(master)
            self._result = None

            self.title("登录 amis.ssjj.cn")
            self.configure(fg_color=CLR["win"])
            self.resizable(False, False)
            self.geometry("470x470")

            f_h1 = ctk.CTkFont(family=FONT_UI, size=18, weight="bold")
            f_key = ctk.CTkFont(family=FONT_UI, size=12)
            f_sub = ctk.CTkFont(family=FONT_UI, size=11)
            f_entry = ctk.CTkFont(family="Consolas", size=14)
            f_btn = ctk.CTkFont(family=FONT_UI, size=13, weight="bold")

            head = ctk.CTkFrame(self, fg_color="transparent")
            head.pack(fill="x", padx=24, pady=(20, 4))
            ctk.CTkLabel(head, text="登录 amis.ssjj.cn", font=f_h1,
                         text_color=CLR["text"]).pack(anchor="w")
            ctk.CTkLabel(head, text=tip or "自动获取 Token 失败，请手动登录后继续抓取。",
                         font=f_sub, text_color=CLR["sub"], wraplength=410,
                         justify="left").pack(anchor="w", pady=(4, 0))

            card = ctk.CTkFrame(self, fg_color=CLR["card"], corner_radius=14,
                                border_width=1, border_color=CLR["line"])
            card.pack(fill="both", expand=True, padx=24, pady=(12, 0))
            card.grid_columnconfigure(1, weight=1)

            # 登录网址（明确展示）
            ctk.CTkLabel(card, text="登录网址", font=f_key,
                         text_color=CLR["sub"]).grid(row=0, column=0, sticky="w",
                                                     padx=(20, 12), pady=(16, 4))
            url_row = ctk.CTkFrame(card, fg_color="transparent")
            url_row.grid(row=0, column=1, sticky="ew", padx=(0, 16), pady=(16, 4))
            url_row.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(url_row, text=LOGIN_PAGE, font=f_entry,
                         text_color=CLR["accent"]).grid(row=0, column=0, sticky="w")
            ctk.CTkButton(url_row, text="打开", width=56, height=28, corner_radius=8,
                          font=f_sub, fg_color=CLR["ghost"], hover_color=CLR["ghost_h"],
                          text_color=CLR["text"],
                          command=self._open_login_page).grid(row=0, column=1, sticky="e")

            # 账号
            ctk.CTkLabel(card, text="账号", font=f_key,
                         text_color=CLR["sub"]).grid(row=1, column=0, sticky="w",
                                                     padx=(20, 12), pady=(10, 0))
            self.ent_account = ctk.CTkEntry(card, height=36, font=f_entry, corner_radius=9,
                                            fg_color=CLR["log_bg"], border_color=CLR["line"],
                                            text_color=CLR["text"], placeholder_text="请输入账号")
            self.ent_account.grid(row=1, column=1, sticky="ew", padx=(0, 16), pady=(10, 0))

            # 密码（掩码显示）
            ctk.CTkLabel(card, text="密码", font=f_key,
                         text_color=CLR["sub"]).grid(row=2, column=0, sticky="w",
                                                     padx=(20, 12), pady=(10, 0))
            self.ent_password = ctk.CTkEntry(card, height=36, font=f_entry, corner_radius=9,
                                             fg_color=CLR["log_bg"], border_color=CLR["line"],
                                             text_color=CLR["text"], show="●",
                                             placeholder_text="请输入密码")
            self.ent_password.grid(row=2, column=1, sticky="ew", padx=(0, 16), pady=(10, 0))

            # LDAP 登录（默认不勾选）
            self.var_ldap = ctk.BooleanVar(value=False)
            ctk.CTkCheckBox(
                card, text="LDAP 登录（本后台使用 LDAP 账号时请勾选）", variable=self.var_ldap,
                font=f_sub, text_color=CLR["text"], checkbox_width=18, checkbox_height=18,
                corner_radius=5, fg_color=CLR["accent"], hover_color=CLR["accent_h"],
                border_color=CLR["line"], checkmark_color="#FFFFFF",
            ).grid(row=3, column=0, columnspan=2, sticky="w", padx=20, pady=(14, 4))

            # 记住账号密码（本机加密保存，默认勾选；登录成功且勾选才写入 DPAPI 加密凭据）
            self.var_remember = ctk.BooleanVar(value=True)
            ctk.CTkCheckBox(
                card, text="记住账号密码（本机加密保存）", variable=self.var_remember,
                font=f_sub, text_color=CLR["text"], checkbox_width=18, checkbox_height=18,
                corner_radius=5, fg_color=CLR["accent"], hover_color=CLR["accent_h"],
                border_color=CLR["line"], checkmark_color="#FFFFFF",
            ).grid(row=4, column=0, columnspan=2, sticky="w", padx=20, pady=(4, 2))

            # 服务端错误提示（重试时回显上一次失败原因）
            self.lbl_error = ctk.CTkLabel(card, text=error or "", font=f_sub,
                                          text_color=ERR_CLR, wraplength=390, justify="left")
            self.lbl_error.grid(row=5, column=0, columnspan=2, sticky="w",
                                padx=20, pady=(6, 16))

            foot = ctk.CTkFrame(self, fg_color="transparent")
            foot.pack(fill="x", padx=24, pady=16)
            ctk.CTkButton(foot, text="登录", width=140, height=40, corner_radius=10,
                          font=f_btn, fg_color=CLR["accent"], hover_color=CLR["accent_h"],
                          text_color="#FFFFFF",
                          command=self._submit).pack(side="left")
            ctk.CTkButton(foot, text="取消", width=100, height=40, corner_radius=10,
                          font=f_btn, fg_color=CLR["ghost"], hover_color=CLR["ghost_h"],
                          text_color=CLR["text"],
                          command=self._cancel).pack(side="right")

            self.protocol("WM_DELETE_WINDOW", self._cancel)
            self.bind("<Return>", lambda _e: self._submit())
            self.bind("<Escape>", lambda _e: self._cancel())
            if default_account:
                self.ent_account.insert(0, default_account)
            self.transient(master)
            self.after(120, self._place)
            self.lift()
            try:
                self.grab_set()
            except Exception:
                pass
            (self.ent_password if default_account else self.ent_account).focus_set()

        def _place(self):
            try:
                self.update_idletasks()
                m = self.master
                x = m.winfo_rootx() + (m.winfo_width() - self.winfo_width()) // 2
                y = m.winfo_rooty() + (m.winfo_height() - self.winfo_height()) // 3
                self.geometry(f"+{max(0, x)}+{max(0, y)}")
            except Exception:
                pass

        def _open_login_page(self):
            try:
                webbrowser.open(LOGIN_PAGE)
            except Exception as e:
                log(f"[警告] 打开登录页失败: {e}")

        def _submit(self):
            account = self.ent_account.get().strip()
            password = self.ent_password.get()
            if not account:
                self.lbl_error.configure(text="请输入账号。")
                return
            if not password:
                self.lbl_error.configure(text="请输入密码。")
                return
            # 凭据仅放入内存结果，不做任何持久化；是否记住由登录成功后统一处理
            self._result = (account, password, bool(self.var_ldap.get()),
                            bool(self.var_remember.get()))
            self.destroy()

        def _cancel(self):
            self._result = None
            self.destroy()

        def wait(self):
            """阻塞至对话框关闭，返回 (account, password, ldap, remember) 或 None（取消）"""
            self.wait_window(self)
            return self._result

    class UpdateDialog(ctk.CTkToplevel):
        """发现新版本：展示版本信息并支持在线更新"""

        def __init__(self, master, info):
            super().__init__(master)
            self.app = master
            self.info = info
            self._busy = False
            self._line = 1        # 当前下载线路序号（>1 时会显示「第 n/总 条线路」）
            self._lines = 1

            self.title("发现新版本")
            self.configure(fg_color=CLR["win"])
            self.resizable(False, False)
            self.geometry("620x480")

            f_title = ctk.CTkFont(family=FONT_UI, size=19, weight="bold")
            f_key = ctk.CTkFont(family=FONT_UI, size=12)
            f_val = ctk.CTkFont(family=FONT_UI, size=13, weight="bold")
            f_body = ctk.CTkFont(family=FONT_UI, size=12)
            f_mono = ctk.CTkFont(family="Consolas", size=11)

            head = ctk.CTkFrame(self, fg_color="transparent")
            head.pack(fill="x", padx=24, pady=(22, 6))
            ctk.CTkLabel(head, text="↑  发现新版本", font=f_title,
                         text_color=CLR["accent"]).pack(side="left")

            body = ctk.CTkFrame(self, fg_color=CLR["card"], corner_radius=14,
                                border_width=1, border_color=CLR["line"])
            body.pack(fill="both", expand=True, padx=24, pady=(10, 0))
            body.grid_columnconfigure(1, weight=1)

            size = info.get("size") or 0
            rows = [
                ("当前版本", f"v{current_app_version()}"),
                ("最新版本", f"v{info.get('version', '')}"),
                ("文件大小", f"{size / 1048576:.1f} MB" if size else "未知"),
            ]
            for i, (k, v) in enumerate(rows):
                ctk.CTkLabel(body, text=k, font=f_key, text_color=CLR["sub"]).grid(
                    row=i, column=0, sticky="w", padx=(20, 16), pady=(14 if i == 0 else 7, 0))
                ctk.CTkLabel(body, text=v, font=f_val, text_color=CLR["text"],
                             anchor="w").grid(row=i, column=1, sticky="w",
                                              padx=(0, 20), pady=(14 if i == 0 else 7, 0))

            # 更新说明先做清洗：界面上不展示任何仓库地址 / 平台名 / 下载引导
            notes = scrub_update_text(info.get("notes") or "")
            ctk.CTkLabel(body, text="更新说明", font=f_key, text_color=CLR["sub"]).grid(
                row=len(rows), column=0, columnspan=2, sticky="w", padx=20, pady=(18, 4))
            txt = ctk.CTkTextbox(body, height=126, corner_radius=10, font=f_mono, wrap="word",
                                 fg_color=CLR["log_bg"], border_width=1,
                                 border_color=CLR["line"], text_color=CLR["text"])
            txt.grid(row=len(rows) + 1, column=0, columnspan=2, sticky="nsew",
                     padx=20, pady=(0, 16))
            txt.insert("1.0", notes[:4000] if notes else "（该版本未提供更新说明）")
            txt.configure(state="disabled")
            body.grid_rowconfigure(len(rows) + 1, weight=1)

            self.lbl_msg = ctk.CTkLabel(self, text="", font=f_body, text_color=CLR["sub"])
            self.lbl_msg.pack(fill="x", padx=26, pady=(10, 0))
            self.progress = ctk.CTkProgressBar(self, height=6, corner_radius=3,
                                               progress_color=CLR["accent"],
                                               fg_color=CLR["ghost"])
            self.progress.set(0)
            self.progress.pack(fill="x", padx=26, pady=(6, 0))

            foot = ctk.CTkFrame(self, fg_color="transparent")
            foot.pack(fill="x", padx=24, pady=18)
            self.btn_go = ctk.CTkButton(
                foot, text="立即更新", width=150, height=40, corner_radius=10,
                font=ctk.CTkFont(family=FONT_UI, size=13, weight="bold"),
                fg_color=CLR["accent"], hover_color=CLR["accent_h"], text_color="#FFFFFF",
                command=self._start_update)
            self.btn_go.pack(side="left")
            self.btn_later = ctk.CTkButton(
                foot, text="稍后", width=100, height=40, corner_radius=10, font=f_body,
                fg_color=CLR["ghost"], hover_color=CLR["ghost_h"], text_color=CLR["text"],
                command=self._close)
            self.btn_later.pack(side="right")

            self.transient(master)
            self.after(120, self._place)
            self.lift()
            try:
                self.grab_set()
            except Exception:
                pass
            self.protocol("WM_DELETE_WINDOW", self._close)

        def _place(self):
            try:
                self.update_idletasks()
                m = self.master
                x = m.winfo_rootx() + (m.winfo_width() - self.winfo_width()) // 2
                y = m.winfo_rooty() + (m.winfo_height() - self.winfo_height()) // 3
                self.geometry(f"+{max(0, x)}+{max(0, y)}")
            except Exception:
                pass

        def _close(self):
            if self._busy:
                return          # 下载中不允许关闭，避免留下半成品
            self.destroy()

        # ---- 下载与替换 ----
        def _start_update(self):
            if self._busy:
                return
            if not current_exe_path():
                # 源码运行：不做自我替换；也不跳转外部页面（不引导用户去仓库取包）
                self._busy = True
                self.btn_go.configure(state="disabled", text="无法自动更新")
                self.lbl_msg.configure(text="当前为源码运行，暂不支持在线更新。")
                return
            self._busy = True
            self.btn_go.configure(state="disabled", text="正在下载…")
            self.btn_later.configure(state="disabled")
            self.lbl_msg.configure(text="正在下载新版本，请勿关闭窗口…")
            threading.Thread(target=self._download_worker, daemon=True).start()

        def _download_worker(self):
            try:
                path = download_update(
                    self.info,
                    on_progress=lambda got, total: self.app._q.put(
                        ("update_progress", (got, total))),
                    on_stage=lambda idx, n: self.app._q.put(("update_stage", (idx, n))),
                )
            except Exception as e:
                self.app._q.put(("update_error", f"{type(e).__name__}: {e}"))
                return
            self.app._q.put(("update_ready", path))

        def _on_stage(self, idx, lines):
            """换到第 idx 条下载线路：记下行号，状态行给一句中性提示。

            文案里不含任何地址 / 平台名，只告诉用户「正在自动换条路重试」。
            """
            try:
                self._line, self._lines = idx, lines
                self.lbl_msg.configure(
                    text=f"当前线路不顺畅，正在自动切换（第 {idx}/{lines} 条）…")
            except Exception:
                pass

        def _on_progress(self, got, total):
            try:
                if total:
                    self.progress.set(min(1.0, got / total))
                    text = f"正在下载… {got / 1048576:.1f} / {total / 1048576:.1f} MB"
                else:
                    text = f"正在下载… {got / 1048576:.1f} MB"
                if self._line > 1:
                    text += f"（第 {self._line}/{self._lines} 条线路）"
                self.lbl_msg.configure(text=text)
            except Exception:
                pass

        def _on_error(self, msg):
            self._busy = False
            # 失败原因也过一遍清洗：日志区、状态行与弹窗里都不出现仓库地址 / 下载引导
            msg = scrub_update_text(msg) or "下载未完成"
            try:
                self.btn_go.configure(state="normal", text="重试")
                self.btn_later.configure(state="normal")
                self.lbl_msg.configure(text=f"更新失败：{msg}")
            except Exception:
                pass
            try:
                import tkinter.messagebox as mb
                mb.showerror("更新失败",
                             f"下载新版本失败：\n\n{msg}\n\n"
                             "请检查网络后点击「重试」，或稍后再试。")
            except Exception:
                pass

        def _on_ready(self, path):
            try:
                self.progress.set(1)
                self.lbl_msg.configure(text="下载完成，正在重启到新版本…")
                old = current_exe_path()
                if old and os.path.abspath(old) != os.path.abspath(path):
                    prepare_update_cleanup(old)
                launch_detached(path)
            except Exception as e:
                self._busy = False
                try:
                    self.btn_go.configure(state="normal", text="重试")
                    self.lbl_msg.configure(text=f"启动新版本失败：{e}，请手动运行 {path}")
                except Exception:
                    pass
                return
            log(f"[更新] 已启动新版本，当前程序即将退出: {path}")
            self.app.after(700, self.app._quit_for_update)

    class DatePickerPopup(ctk.CTkToplevel):
        """日期选择弹层（自绘月历，不引入第三方 datepicker）"""

        WEEK = ("一", "二", "三", "四", "五", "六", "日")

        def __init__(self, master, anchor, day="", on_pick=None):
            super().__init__(master)
            self._on_pick = on_pick
            try:
                self._sel = datetime.strptime(day, "%Y%m%d")
            except Exception:
                self._sel = datetime.now()
            self._shown = self._sel.replace(day=1)

            self.overrideredirect(True)          # 无标题栏，像一个下拉层
            self.configure(fg_color=CLR["card"])
            self.attributes("-topmost", True)

            self.f_small = ctk.CTkFont(family=FONT_UI, size=11)
            self.f_month = ctk.CTkFont(family=FONT_UI, size=13, weight="bold")

            head = ctk.CTkFrame(self, fg_color="transparent")
            head.pack(fill="x", padx=10, pady=(10, 4))
            ctk.CTkButton(head, text="‹", width=26, height=26, corner_radius=7,
                          font=self.f_small, fg_color="transparent", hover_color=CLR["ghost"],
                          text_color=CLR["text"], command=lambda: self._shift(-1)).pack(side="left")
            self.lbl_month = ctk.CTkLabel(head, text="", font=self.f_month, text_color=CLR["text"])
            self.lbl_month.pack(side="left", expand=True)
            ctk.CTkButton(head, text="›", width=26, height=26, corner_radius=7,
                          font=self.f_small, fg_color="transparent", hover_color=CLR["ghost"],
                          text_color=CLR["text"], command=lambda: self._shift(1)).pack(side="right")

            self.grid_box = ctk.CTkFrame(self, fg_color="transparent")
            self.grid_box.pack(padx=10, pady=(0, 6))

            foot = ctk.CTkFrame(self, fg_color="transparent")
            foot.pack(fill="x", padx=10, pady=(0, 10))
            ctk.CTkButton(foot, text="今天", width=64, height=28, corner_radius=8,
                          font=self.f_small, fg_color=CLR["ghost"], hover_color=CLR["ghost_h"],
                          text_color=CLR["text"],
                          command=lambda: self._pick(date.today())).pack(side="left")
            ctk.CTkButton(foot, text="取消", width=56, height=28, corner_radius=8,
                          font=self.f_small, fg_color="transparent", hover_color=CLR["ghost"],
                          text_color=CLR["sub"], command=self.destroy).pack(side="right")

            self._render()
            self.bind("<Escape>", lambda _e: self.destroy())
            self.after(10, lambda: self._place(anchor))
            self.after(20, self._grab)

        def _grab(self):
            try:
                self.grab_set()
            except Exception:
                pass
            self.focus_force()

        def _place(self, anchor):
            try:
                self.update_idletasks()
                x = anchor.winfo_rootx()
                y = anchor.winfo_rooty() + anchor.winfo_height() + 6
                sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
                w, h = self.winfo_reqwidth(), self.winfo_reqheight()
                x = max(0, min(x, sw - w - 8))
                if y + h > sh:
                    y = anchor.winfo_rooty() - h - 6
                self.geometry(f"+{x}+{y}")
            except Exception:
                pass

        def _shift(self, delta):
            y, m = self._shown.year, self._shown.month + delta
            if m < 1:
                y, m = y - 1, 12
            elif m > 12:
                y, m = y + 1, 1
            self._shown = self._shown.replace(year=y, month=m)
            self._render()

        def _pick(self, d):
            try:
                if self._on_pick:
                    self._on_pick(d.strftime("%Y%m%d"))
            finally:
                self.destroy()

        def _render(self):
            for w in self.grid_box.winfo_children():
                w.destroy()
            y, m = self._shown.year, self._shown.month
            self.lbl_month.configure(text=f"{y} 年 {m} 月")
            for i, wd in enumerate(self.WEEK):
                ctk.CTkLabel(self.grid_box, text=wd, width=34, font=self.f_small,
                             text_color=CLR["sub"]).grid(row=0, column=i, padx=1, pady=(0, 2))
            today = date.today()
            for r, week in enumerate(calendar.monthcalendar(y, m), start=1):
                for c, d in enumerate(week):
                    if not d:
                        continue
                    is_sel = (d == self._sel.day and m == self._sel.month and y == self._sel.year)
                    is_today = (d == today.day and m == today.month and y == today.year)
                    ctk.CTkButton(
                        self.grid_box, text=str(d), width=34, height=28, corner_radius=7,
                        font=self.f_small,
                        fg_color=CLR["accent"] if is_sel else ("transparent"),
                        hover_color=CLR["accent_h"] if is_sel else CLR["ghost"],
                        text_color="#FFFFFF" if is_sel else (CLR["accent"] if is_today else CLR["text"]),
                        border_width=0 if (is_sel or not is_today) else 1,
                        border_color=CLR["accent"],
                        command=lambda dd=d: self._pick(date(y, m, dd)),
                    ).grid(row=r, column=c, padx=1, pady=1)

    class ReplayApp(ctk.CTk):
        """主窗口：左侧导航 + 右侧「工作台 / 历史记录 / 设置」三个页面"""

        STAGE_ORDER = ("fetch", "csv", "process", "clip")
        STAGE_TEXT = {"fetch": "抓取", "csv": "导出", "process": "处理", "clip": "复制"}

        def __init__(self, selfcheck=False):
            global _LOG_SINK, _LOGIN_UI
            super().__init__()

            # 窗口图标：iconbitmap(default=...) 一次设好，所有 Toplevel（结果窗 / 登录 / 更新）自动继承。
            # 打包成 exe 后 .ico 在 sys._MEIPASS/assets 下，源码运行时在脚本目录下。
            _ico = app_icon_path()
            if _ico:
                try:
                    self.iconbitmap(default=_ico)
                except Exception:
                    pass

            # 授权校验出结果之前先把窗口藏起来（否则网络慢时界面会先露出来几十秒）
            self._ui_hidden = bool(LICENSE_ENABLED) and not selfcheck
            if self._ui_hidden:
                self.withdraw()

            self._q = queue.Queue()
            self._running = False
            self._update_dialog = None
            self._cfg = load_settings(force=True)
            self._page = "work"
            self._nav = {}
            self._pages = {}
            self._preview_rows = []
            self._preview_range = self._cfg.get("clipboard_range") or CLIPBOARD_RANGE
            self._last_out = ""
            self._hist_rows = {}
            self._last_day = ""

            # 动效状态：_tweens 存各条补间动画的 after id，同一 key 后来的会顶掉先前的
            self._tweens = {}
            self._progress_val = 0.0
            self._pill_busy = False

            self._day_default = (date.today() - timedelta(days=1)).strftime("%Y%m%d")

            self._license_ok = not LICENSE_ENABLED
            self._license_pending = bool(LICENSE_ENABLED)
            self._license_info = None
            self._license_guard = None
            self._pending_day = None

            ctk.set_appearance_mode("dark" if self._cfg.get("theme") == "dark" else "light")
            self.title(f"{APP_TITLE}  v{current_app_version()}")
            self.configure(fg_color=CLR["win"])
            self.geometry(self._center_geometry(1180, 800))
            self.minsize(1080, 720)

            self.f_h1 = ctk.CTkFont(family=FONT_UI, size=22, weight="bold")
            self.f_h2 = ctk.CTkFont(family=FONT_UI, size=15, weight="bold")
            self.f_sec = ctk.CTkFont(family=FONT_UI, size=13, weight="bold")
            self.f_sub = ctk.CTkFont(family=FONT_UI, size=12)
            self.f_small = ctk.CTkFont(family=FONT_UI, size=11)
            self.f_body = ctk.CTkFont(family=FONT_UI, size=12)
            self.f_btn = ctk.CTkFont(family=FONT_UI, size=13, weight="bold")
            self.f_num = ctk.CTkFont(family=FONT_UI, size=20, weight="bold")
            self.f_entry = ctk.CTkFont(family="Consolas", size=13)
            self.f_pv = ctk.CTkFont(family=FONT_UI, size=11)
            self.f_log = ctk.CTkFont(family="Consolas", size=11)

            self.grid_columnconfigure(1, weight=1)
            self.grid_rowconfigure(0, weight=1)

            self._build_sidebar()
            self._build_body()
            self._show_page("work")
            self._refresh_recent()
            self._render_preview([], self._preview_range)

            self.protocol("WM_DELETE_WINDOW", self._on_close)
            _LOG_SINK = lambda text: self._q.put(("log", text))
            _LOGIN_UI = self._ask_login_threadsafe
            self._append_log(f"{datetime.now():%H:%M:%S}  界面已就绪，默认抓取日期 {self._day_default}")
            self.after(80, self._drain)
            threading.Thread(target=self._preheat, daemon=True).start()
            if not selfcheck:
                self.after(120, self._start_license_check)

        # ---------------- 基础 ----------------
        def _center_geometry(self, w, h):
            try:
                sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
                return f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h) // 2 - 30)}"
            except Exception:
                return f"{w}x{h}"

        def _pal(self, key):
            """取当前外观下的具体色值（给 tk 原生控件用）"""
            value = CLR[key]
            try:
                return value[0] if ctk.get_appearance_mode() == "Light" else value[1]
            except Exception:
                return value[0]

        # ---------------- 左侧导航 ----------------
        def _build_sidebar(self):
            bar = ctk.CTkFrame(self, width=216, corner_radius=0, fg_color=CLR["side"])
            bar.grid(row=0, column=0, sticky="nsw")
            bar.grid_propagate(False)
            bar.pack_propagate(False)      # 子控件是 pack 的，必须禁 pack 传播才是固定宽度

            brand = ctk.CTkFrame(bar, fg_color="transparent")
            brand.pack(fill="x", padx=18, pady=(20, 16))
            logo = ctk.CTkFrame(brand, width=34, height=34, corner_radius=10, fg_color=CLR["accent"])
            logo.pack(side="left")
            logo.pack_propagate(False)
            ctk.CTkLabel(logo, text="AI", font=ctk.CTkFont(family=FONT_UI, size=12, weight="bold"),
                         text_color="#FFFFFF").pack(expand=True)
            box = ctk.CTkFrame(brand, fg_color="transparent")
            box.pack(side="left", padx=(10, 0))
            ctk.CTkLabel(box, text="AIReplay", font=self.f_h2, text_color=CLR["text"]).pack(anchor="w")
            ctk.CTkLabel(box, text="Replay 核查", font=self.f_small,
                         text_color=CLR["sub"]).pack(anchor="w")

            for key, text in (("work", "工作台"), ("history", "历史记录"), ("settings", "设置")):
                self._nav_item(bar, key, text)

            ctk.CTkFrame(bar, height=1, fg_color=CLR["line"]).pack(fill="x", padx=18, pady=(14, 10))
            ctk.CTkLabel(bar, text="最近任务", font=self.f_small,
                         text_color=CLR["sub"]).pack(anchor="w", padx=20)
            self.recent_box = ctk.CTkFrame(bar, fg_color="transparent")
            self.recent_box.pack(fill="x", padx=14, pady=(4, 0))

            bottom = ctk.CTkFrame(bar, fg_color="transparent")
            bottom.pack(side="bottom", fill="x", padx=16, pady=(0, 16))
            self.seg_theme = ctk.CTkSegmentedButton(
                bottom, values=["深色", "浅色"], height=28, font=self.f_small,
                command=self._on_theme_change, selected_color=CLR["accent"],
                selected_hover_color=CLR["accent_h"], unselected_color=CLR["ghost"],
                unselected_hover_color=CLR["ghost_h"], text_color=CLR["text"],
            )
            self.seg_theme.set("深色" if self._cfg.get("theme") == "dark" else "浅色")
            self.seg_theme.pack(fill="x")
            self.btn_check = ctk.CTkButton(
                bottom, text="检查更新", height=28, corner_radius=8, font=self.f_small,
                fg_color="transparent", hover_color=CLR["ghost"], text_color=CLR["sub"],
                command=lambda: self._check_update(manual=True))
            self.btn_check.pack(fill="x", pady=(8, 4))
            ctk.CTkLabel(bottom, text=f"v{current_app_version()}", font=self.f_small,
                         text_color=CLR["sub"]).pack(anchor="w")

        def _nav_item(self, parent, key, text):
            row = ctk.CTkFrame(parent, height=38, corner_radius=9, fg_color="transparent")
            row.pack(fill="x", padx=12, pady=2)
            row.pack_propagate(False)
            mark = ctk.CTkFrame(row, width=3, height=18, corner_radius=2, fg_color="transparent")
            mark.pack(side="left", padx=(8, 10))
            lbl = ctk.CTkLabel(row, text=text, font=self.f_body, text_color=CLR["text"], anchor="w")
            lbl.pack(side="left", fill="both", expand=True)
            for w in (row, lbl):
                w.bind("<Button-1>", lambda _e, k=key: self._show_page(k))
                w.bind("<Enter>", lambda _e, r=row, k=key: self._nav_hover(r, k, True))
                w.bind("<Leave>", lambda _e, r=row, k=key: self._nav_hover(r, k, False))
            self._nav[key] = (row, mark, lbl)

        def _nav_hover(self, row, key, entering):
            if self._page == key:
                return
            row.configure(fg_color=CLR["nav_hover"] if entering else "transparent")

        def _paint_nav(self):
            for key, (row, mark, lbl) in self._nav.items():
                on = key == self._page
                row.configure(fg_color=CLR["nav_on"] if on else "transparent")
                mark.configure(fg_color=CLR["accent"] if on else "transparent")
                lbl.configure(text_color=CLR["accent"] if on else CLR["text"])

        def _show_page(self, key):
            """切页：三页常驻叠放，只靠 lift 换前后景。

            不 unmap/map（grid_remove + grid 重新映射一个百来控件的页面实测 1.6~2.0 秒），
            也不做位移补间 —— 切页就是瞬时切换，只有数据变化时才有动效。"""
            self._page = key
            if key == "history":
                self._refresh_history()
            elif key == "settings":
                self._load_settings_into_form()

            # 三页在 _build_body 里就已经 place 满了整个 body，这里只换前后景，
            # 不要重复 place：对一个已就位的页面再 place 一次会让整个页面重排（实测 20ms+）
            self._pages[key].lift()
            self._paint_nav()

        # ---------------- 页面容器 ----------------
        def _build_body(self):
            self.body = ctk.CTkFrame(self, fg_color="transparent")
            self.body.grid(row=0, column=1, sticky="nsew")
            self.body.grid_columnconfigure(0, weight=1)
            self.body.grid_rowconfigure(0, weight=1)
            for key, builder in (("work", self._build_work_page),
                                 ("history", self._build_history_page),
                                 ("settings", self._build_settings_page)):
                page = builder()
                page.place(x=0, y=0, relwidth=1, relheight=1)
                self._pages[key] = page

        # ---------------- 工作台 ----------------
        def _build_work_page(self):
            # 三个页面常驻显示、靠 lift 叠放，切换时不做 map/unmap：
            # 每次重新映射这一百多个控件要花上百毫秒，切页会明显卡一下
            page = ctk.CTkFrame(self.body, fg_color=CLR["win"])
            page.grid_columnconfigure(0, weight=1)
            page.grid_rowconfigure(4, weight=1)      # 预览区吃掉剩余高度

            head = ctk.CTkFrame(page, fg_color="transparent")
            head.grid(row=0, column=0, sticky="ew", padx=24, pady=(20, 12))
            head.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(head, text="核查工作台", font=self.f_h1,
                         text_color=CLR["text"]).grid(row=0, column=0, sticky="w")
            ctk.CTkLabel(head, text="选日期 → 抓取 → 导出 → 去重 → 复制到剪贴板",
                         font=self.f_sub, text_color=CLR["sub"]).grid(row=1, column=0, sticky="w", pady=(3, 0))
            self.lbl_status = ctk.CTkLabel(head, text="●  就绪", font=self.f_small, width=96, height=28,
                                           corner_radius=14, fg_color=CLR["ghost"], text_color=CLR["sub"])
            self.lbl_status.grid(row=0, column=1, rowspan=2, sticky="e")

            # 任务条
            task = ctk.CTkFrame(page, fg_color=CLR["card"], corner_radius=14,
                                border_width=1, border_color=CLR["line"])
            task.grid(row=1, column=0, sticky="ew", padx=24)
            row = ctk.CTkFrame(task, fg_color="transparent")
            row.pack(fill="x", padx=16, pady=(14, 8))
            ctk.CTkLabel(row, text="抓取日期", font=self.f_body,
                         text_color=CLR["sub"]).pack(side="left", padx=(0, 10))
            self.entry_day = ctk.CTkEntry(row, width=118, height=34, font=self.f_entry,
                                          corner_radius=8, justify="center",
                                          fg_color=CLR["log_bg"], border_color=CLR["line"],
                                          text_color=CLR["text"])
            self.entry_day.pack(side="left")
            self.entry_day.insert(0, self._day_default)
            for text, off in (("昨天", 1), ("今天", 0), ("前天", 2)):
                ctk.CTkButton(row, text=text, width=52, height=34, corner_radius=8, font=self.f_small,
                              fg_color=CLR["ghost"], hover_color=CLR["ghost_h"], text_color=CLR["text"],
                              command=lambda o=off: self._set_quick_day(o)).pack(side="left", padx=(8, 0))
            self.btn_cal = ctk.CTkButton(row, text="日历", width=56, height=34, corner_radius=8,
                                         font=self.f_small, fg_color=CLR["ghost"],
                                         hover_color=CLR["ghost_h"], text_color=CLR["text"],
                                         command=self._open_calendar)
            self.btn_cal.pack(side="left", padx=(8, 0))

            self.btn_open = ctk.CTkButton(row, text="打开目录", width=96, height=34, corner_radius=8,
                                          font=self.f_small, fg_color=CLR["ghost"],
                                          hover_color=CLR["ghost_h"], text_color=CLR["text"],
                                          command=self._open_output_folder)
            self.btn_open.pack(side="right")
            self.btn_start = ctk.CTkButton(row, text="开始核查", width=126, height=38, corner_radius=9,
                                           font=self.f_btn, fg_color=CLR["accent"],
                                           hover_color=CLR["accent_h"], text_color="#FFFFFF",
                                           command=self._on_start)
            self.btn_start.pack(side="right", padx=(0, 10))

            stages = ctk.CTkFrame(task, fg_color="transparent")
            stages.pack(fill="x", padx=16, pady=(0, 8))
            self._stage_lbl = {}
            for i, key in enumerate(self.STAGE_ORDER):
                if i:
                    ctk.CTkLabel(stages, text="›", font=self.f_small,
                                 text_color=CLR["sub"]).pack(side="left", padx=6)
                lbl = ctk.CTkLabel(stages, text=self.STAGE_TEXT[key], font=self.f_small,
                                   text_color=CLR["sub"])
                lbl.pack(side="left")
                self._stage_lbl[key] = lbl
            self.lbl_stage_note = ctk.CTkLabel(stages, text="", font=self.f_small,
                                               text_color=CLR["sub"])
            self.lbl_stage_note.pack(side="right")

            # 细进度条：空闲时静止，运行时来回滚动（indeterminate）
            self.progress = ctk.CTkProgressBar(task, height=4, corner_radius=2,
                                               fg_color=CLR["ghost"],
                                               progress_color=CLR["accent"])
            self.progress.set(0)
            self.progress.pack(fill="x", padx=16, pady=(0, 12))

            # 统计
            tiles = ctk.CTkFrame(page, fg_color="transparent")
            tiles.grid(row=2, column=0, sticky="ew", padx=24, pady=(12, 0))
            self._stat = {}
            for i, (key, label) in enumerate((("fetch", "抓取条数"), ("filter", "筛选后"),
                                              ("dedup", "去重后"))):
                tiles.grid_columnconfigure(i, weight=1)
                tile = ctk.CTkFrame(tiles, fg_color=CLR["tile"], corner_radius=12)
                tile.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 10, 0))
                ctk.CTkLabel(tile, text=label, font=self.f_small,
                             text_color=CLR["sub"]).pack(anchor="w", padx=16, pady=(11, 0))
                value = ctk.CTkLabel(tile, text="—", font=self.f_num, text_color=CLR["text"])
                value.pack(anchor="w", padx=16, pady=(0, 10))
                self._stat[key] = value

            # 已复制提示条（成功后出现）
            self.clip_bar = ctk.CTkFrame(page, fg_color=CLR["ok_bg"], corner_radius=10,
                                         border_width=1, border_color=CLR["ok_line"])
            self.clip_bar.grid(row=3, column=0, sticky="ew", padx=24, pady=(10, 0))
            self.clip_lbl = ctk.CTkLabel(self.clip_bar, text="", font=self.f_body, text_color=CLR["ok_text"])
            self.clip_lbl.pack(side="left", padx=16, pady=9)
            ctk.CTkButton(self.clip_bar, text="再次复制", width=84, height=26, corner_radius=7,
                          font=self.f_small, fg_color="transparent", border_width=1,
                          border_color=CLR["ok_line"], hover_color=CLR["ok_bg"],
                          text_color=CLR["ok_text"], command=self._copy_again).pack(side="right", padx=10)
            self.clip_bar.grid_remove()

            self._build_preview(page)
            self._build_log(page)
            return page

        def _build_preview(self, page):
            card = ctk.CTkFrame(page, fg_color=CLR["card"], corner_radius=14,
                                border_width=1, border_color=CLR["line"])
            card.grid(row=4, column=0, sticky="nsew", padx=24, pady=(12, 0))
            card.grid_columnconfigure(0, weight=1)
            card.grid_rowconfigure(1, weight=1)

            head = ctk.CTkFrame(card, fg_color="transparent")
            head.grid(row=0, column=0, sticky="ew", padx=16, pady=(10, 6))
            head.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(head, text="复制内容预览", font=self.f_sec,
                         text_color=CLR["text"]).grid(row=0, column=0, sticky="w")
            self.lbl_range = ctk.CTkLabel(head, text=self._preview_range, font=self.f_small,
                                          text_color=CLR["sub"])
            self.lbl_range.grid(row=0, column=1, sticky="e")

            holder = ctk.CTkFrame(card, fg_color=CLR["log_bg"], corner_radius=10)
            holder.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
            holder.grid_columnconfigure(0, weight=1)
            holder.grid_rowconfigure(0, weight=1)
            self._pv_holder = holder
            self._mount_preview_canvas()

        # 预览表是「只读的一屏文字」，不需要 110 个 CTkLabel 当控件：
        # 直接在 canvas 上画文字 + 隔行色块，渲染和重绘都快一个数量级
        # （用 110 个 CTkLabel 时，切回工作台要重绘这一屏，单帧 180ms，界面会顿一下）
        PV_COL_W = (86, 86, 80, 70, 52, 108, 108, 70, 108, 74)
        PV_ROW_H = 25
        PV_PAD = 10

        def _mount_preview_canvas(self):
            import tkinter as tk
            for w in self._pv_holder.winfo_children():
                w.destroy()
            bg = self._pal("log_bg")
            canvas = tk.Canvas(self._pv_holder, bg=bg, highlightthickness=0, bd=0)
            canvas.grid(row=0, column=0, sticky="nsew")
            vsb = ctk.CTkScrollbar(self._pv_holder, width=12, command=canvas.yview)
            vsb.grid(row=0, column=1, sticky="ns")
            hsb = ctk.CTkScrollbar(self._pv_holder, height=12, orientation="horizontal",
                                   command=canvas.xview)
            hsb.grid(row=1, column=0, sticky="ew")
            canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
            canvas.bind("<MouseWheel>", self._pv_wheel)
            self._pv_canvas = canvas
            self._render_preview(self._preview_rows, self._preview_range)

        def _pv_font(self, bold=False):
            import tkinter.font as tkfont
            fonts = getattr(self, "_pv_fonts", None)
            if fonts is None:
                fonts = self._pv_fonts = {}
            key = "bold" if bold else "normal"
            if key not in fonts:
                fonts[key] = tkfont.Font(family=FONT_UI, size=9,
                                         weight="bold" if bold else "normal")
            return fonts[key]

        @staticmethod
        def _pv_ellipsis(text, max_px, font):
            """超出列宽就截断加省略号（canvas 文字没有控件那样的裁剪）"""
            if max_px <= 0 or not text:
                return ""
            if font.measure(text) <= max_px:
                return text
            lo, hi = 0, len(text)
            while lo < hi:
                mid = (lo + hi + 1) // 2
                if font.measure(text[:mid] + "…") <= max_px:
                    lo = mid
                else:
                    hi = mid - 1
            return (text[:lo] + "…") if lo else "…"

        def _pv_col_x(self, ncol):
            xs, x = [], self.PV_PAD
            for c in range(ncol):
                xs.append(x)
                x += self.PV_COL_W[c] if c < len(self.PV_COL_W) else 96
            return xs, x + self.PV_PAD

        def _pv_wheel(self, event):
            try:
                if event.state & 0x1:
                    self._pv_canvas.xview_scroll(-1 if event.delta > 0 else 1, "units")
                else:
                    self._pv_canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")
            except Exception:
                pass

        def _render_preview(self, rows, cell_range=None):
            self._preview_rows = rows or []
            if cell_range:
                self._preview_range = cell_range
            try:
                self.lbl_range.configure(text=self._preview_range)
            except Exception:
                pass
            cv = getattr(self, "_pv_canvas", None)
            if cv is None:
                return
            cv.delete("all")
            if not self._preview_rows:
                cv.create_text(self.PV_PAD * 2, 26, anchor="nw",
                               text="跑完一次核查后，这里显示被复制走的表格内容",
                               fill=self._pal("sub"), font=self._pv_font())
                cv.configure(scrollregion=(0, 0, 420, 90))
                return

            rows = self._preview_rows
            ncol = max(len(r) for r in rows)
            xs, total_w = self._pv_col_x(ncol)
            total_h = len(rows) * self.PV_ROW_H + 8
            for r in range(1, len(rows)):                 # 隔行底色（表头不着色）
                if r % 2 == 0:
                    cv.create_rectangle(0, r * self.PV_ROW_H, total_w,
                                        (r + 1) * self.PV_ROW_H,
                                        fill=self._pal("stripe"), width=0)
            for r, row in enumerate(rows):
                header = (r == 0)
                font = self._pv_font(bold=header)
                color = self._pal("sub") if header else self._pal("text")
                y = r * self.PV_ROW_H + self.PV_ROW_H // 2
                for c, val in enumerate(row):
                    if not val:
                        continue
                    room = (self.PV_COL_W[c] if c < len(self.PV_COL_W) else 96) - 10
                    cv.create_text(xs[c], y, anchor="w",
                                   text=self._pv_ellipsis(str(val), room, font),
                                   fill=color, font=font)
            cv.configure(scrollregion=(0, 0, total_w, total_h))

        def _build_log(self, page):
            card = ctk.CTkFrame(page, fg_color=CLR["card"], corner_radius=14,
                                border_width=1, border_color=CLR["line"])
            card.grid(row=5, column=0, sticky="ew", padx=24, pady=(12, 18))
            card.grid_columnconfigure(0, weight=1)
            self.log_card = card
            head = ctk.CTkFrame(card, fg_color="transparent")
            head.grid(row=0, column=0, sticky="ew", padx=16, pady=(9, 4))
            head.grid_columnconfigure(0, weight=1)
            self.btn_log = ctk.CTkButton(head, text="运行日志  ▴", width=96, height=24, corner_radius=7,
                                         font=self.f_sec, fg_color="transparent",
                                         hover_color=CLR["ghost"], text_color=CLR["text"],
                                         anchor="w", command=self._toggle_log)
            self.btn_log.grid(row=0, column=0, sticky="w")
            ctk.CTkButton(head, text="复制日志", width=76, height=24, corner_radius=7, font=self.f_small,
                          fg_color="transparent", hover_color=CLR["ghost"], text_color=CLR["sub"],
                          command=self._copy_log).grid(row=0, column=1, padx=(0, 6))
            ctk.CTkButton(head, text="清空", width=52, height=24, corner_radius=7, font=self.f_small,
                          fg_color="transparent", hover_color=CLR["ghost"], text_color=CLR["sub"],
                          command=self._clear_log).grid(row=0, column=2)

            self.txt = ctk.CTkTextbox(card, height=100, corner_radius=10, font=self.f_log, wrap="word",
                                      fg_color=CLR["log_bg"], border_width=1,
                                      border_color=CLR["line"], text_color=CLR["text"])
            self.txt.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 12))
            self.txt.configure(state="disabled")
            self._log_open = True

        # ---------------- 历史记录 ----------------
        def _build_history_page(self):
            # 三个页面常驻显示、靠 lift 叠放，切换时不做 map/unmap：
            # 每次重新映射这一百多个控件要花上百毫秒，切页会明显卡一下
            page = ctk.CTkFrame(self.body, fg_color=CLR["win"])
            page.grid_columnconfigure(0, weight=1)
            page.grid_rowconfigure(1, weight=1)

            head = ctk.CTkFrame(page, fg_color="transparent")
            head.grid(row=0, column=0, sticky="ew", padx=24, pady=(20, 12))
            head.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(head, text="历史记录", font=self.f_h1,
                         text_color=CLR["text"]).grid(row=0, column=0, sticky="w")
            ctk.CTkLabel(head, text="输出目录里的报表，可直接再次复制 C1:L11",
                         font=self.f_sub, text_color=CLR["sub"]).grid(row=1, column=0, sticky="w", pady=(3, 0))
            ctk.CTkButton(head, text="刷新", width=76, height=32, corner_radius=8, font=self.f_small,
                          fg_color=CLR["ghost"], hover_color=CLR["ghost_h"], text_color=CLR["text"],
                          command=self._refresh_history).grid(row=0, column=1, rowspan=2, sticky="e")

            self.hist_box = ctk.CTkFrame(page, fg_color=CLR["card"], corner_radius=14,
                                         border_width=1, border_color=CLR["line"])
            self.hist_box.grid(row=1, column=0, sticky="nsew", padx=24, pady=(0, 18))
            self.hist_box.grid_columnconfigure(0, weight=1)
            self.hist_box.grid_rowconfigure(1, weight=1)

            cols = ctk.CTkFrame(self.hist_box, fg_color="transparent")
            cols.grid(row=0, column=0, sticky="ew", padx=18, pady=(12, 6))
            for text, width in (("日期", 96), ("行数", 70), ("大小", 80), ("修改时间", 150)):
                ctk.CTkLabel(cols, text=text, width=width, anchor="w", font=self.f_small,
                             text_color=CLR["sub"]).pack(side="left")
            ctk.CTkLabel(cols, text="操作", anchor="e", font=self.f_small,
                         text_color=CLR["sub"]).pack(side="right")

            self.hist_list = ctk.CTkScrollableFrame(self.hist_box, fg_color="transparent")
            self.hist_list.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 12))
            return page

        def _refresh_history(self):
            for w in self.hist_list.winfo_children():
                w.destroy()
            items = scan_outputs(limit=self._cfg.get("history_limit", 80))
            if not items:
                ctk.CTkLabel(self.hist_list, text="还没有产物，去工作台跑一次核查吧",
                             font=self.f_body, text_color=CLR["sub"]).pack(anchor="w", padx=16, pady=20)
                return
            self._hist_rows = {}
            for it in items:
                row = ctk.CTkFrame(self.hist_list, fg_color="transparent")
                row.pack(fill="x", padx=6, pady=2)
                ctk.CTkLabel(row, text=f"{it['day'][:4]}-{it['day'][4:6]}-{it['day'][6:]}",
                             width=96, anchor="w", font=self.f_body,
                             text_color=CLR["text"]).pack(side="left")
                lbl_rows = ctk.CTkLabel(row, text="…" if it["rows"] is None else str(it["rows"]),
                                        width=70, anchor="w", font=self.f_small,
                                        text_color=CLR["sub"])
                lbl_rows.pack(side="left")
                self._hist_rows[it["path"]] = lbl_rows
                ctk.CTkLabel(row, text=f"{it['size'] / 1024:.0f} KB", width=80, anchor="w",
                             font=self.f_small, text_color=CLR["sub"]).pack(side="left")
                ctk.CTkLabel(row, text=datetime.fromtimestamp(it["mtime"]).strftime("%m-%d %H:%M"),
                             width=150, anchor="w", font=self.f_small,
                             text_color=CLR["sub"]).pack(side="left")
                ctk.CTkButton(row, text="再次复制", width=80, height=26, corner_radius=7,
                              font=self.f_small, fg_color=CLR["ghost"], hover_color=CLR["ghost_h"],
                              text_color=CLR["text"],
                              command=lambda p=it["path"]: self._copy_for(p)).pack(side="right")
                ctk.CTkButton(row, text="打开目录", width=80, height=26, corner_radius=7,
                              font=self.f_small, fg_color="transparent", hover_color=CLR["ghost"],
                              text_color=CLR["sub"],
                              command=lambda p=os.path.dirname(it["path"]): self._open_path(p)).pack(
                    side="right", padx=(0, 6))
            threading.Thread(target=self._fill_row_counts,
                             args=([i["path"] for i in items if i["rows"] is None],),
                             daemon=True).start()

        def _fill_row_counts(self, paths):
            for path in paths:
                n = count_sheet_rows(path)
                if n is not None:
                    self._q.put(("hist_rows", (path, n)))

        # ---------------- 设置 ----------------
        def _build_settings_page(self):
            # 三个页面常驻显示、靠 lift 叠放，切换时不做 map/unmap：
            # 每次重新映射这一百多个控件要花上百毫秒，切页会明显卡一下
            page = ctk.CTkFrame(self.body, fg_color=CLR["win"])
            page.grid_columnconfigure(0, weight=1)
            page.grid_rowconfigure(1, weight=1)
            head = ctk.CTkFrame(page, fg_color="transparent")
            head.grid(row=0, column=0, sticky="ew", padx=24, pady=(20, 12))
            ctk.CTkLabel(head, text="设置", font=self.f_h1, text_color=CLR["text"]).pack(side="left")
            ctk.CTkLabel(head, text="改完点保存即时生效，无需重新发版", font=self.f_sub,
                         text_color=CLR["sub"]).pack(side="left", padx=(12, 0), pady=(8, 0))

            scroll = ctk.CTkScrollableFrame(page, fg_color="transparent")
            scroll.grid(row=1, column=0, sticky="nsew", padx=22, pady=(0, 16))

            self._set_vars = {}
            body = self._section(scroll, "筛选条件", "只有同时满足以下条件的行才会进入报表")
            cfg_filters = self._cfg.get("filters") or dict(FILTER_COLUMNS)
            for col, val in cfg_filters.items():
                row = self._setting_row(body, col)
                ent = self._entry(row, width=180)
                ent.insert(0, str(val))
                self._set_vars[f"filter::{col}"] = ent

            body = self._section(scroll, "排序", "按哪一列排序，分数列会按数值比较")
            row = self._setting_row(body, "排序列")
            ent = self._entry(row, width=180)
            ent.insert(0, self._cfg.get("sort_keyword", SORT_COLUMN_KEYWORD))
            self._set_vars["sort_keyword"] = ent
            self.seg_sort = ctk.CTkSegmentedButton(row, values=["降序", "升序"], height=28,
                                                   font=self.f_small, selected_color=CLR["accent"],
                                                   selected_hover_color=CLR["accent_h"],
                                                   unselected_color=CLR["ghost"],
                                                   unselected_hover_color=CLR["ghost_h"],
                                                   text_color=CLR["text"])
            self.seg_sort.set("降序" if self._cfg.get("sort_desc", True) else "升序")
            self.seg_sort.pack(side="left", padx=(10, 0))

            body = self._section(scroll, "去重", "同一角色只保留排序后的第一行")
            row = self._setting_row(body, "去重列")
            ent = self._entry(row, width=180)
            ent.insert(0, self._cfg.get("dedup_keyword", DEDUP_COLUMN_KEYWORD))
            self._set_vars["dedup_keyword"] = ent

            body = self._section(scroll, "剪贴板", "报表生成后自动复制的区域（Excel 记法）")
            row = self._setting_row(body, "复制区域")
            ent = self._entry(row, width=140)
            ent.insert(0, self._cfg.get("clipboard_range") or CLIPBOARD_RANGE)
            self._set_vars["clipboard_range"] = ent
            self.var_clip = ctk.BooleanVar(value=bool(self._cfg.get("clipboard_enabled", True)))
            ctk.CTkSwitch(row, text="输出后自动复制", variable=self.var_clip, font=self.f_small,
                          text_color=CLR["text"], progress_color=CLR["accent"],
                          button_color=CLR["card"]).pack(side="left", padx=(14, 0))

            body = self._section(scroll, "完成后", "默认不弹窗，跑完即可直接粘贴")
            self.var_popup = ctk.BooleanVar(value=bool(self._cfg.get("popup_on_done", False)))
            ctk.CTkSwitch(body, text="弹出结果窗口（会打断当前操作）", variable=self.var_popup,
                          font=self.f_small, text_color=CLR["text"],
                          progress_color=CLR["accent"], button_color=CLR["card"]).pack(anchor="w")

            body = self._section(scroll, "界面", "进度条平滑推进、统计数字滚动、运行中状态呼吸")
            self.var_anim = ctk.BooleanVar(value=bool(self._cfg.get("animations", True)))
            ctk.CTkSwitch(body, text="开启动效（关掉则所有状态切换都是瞬时）",
                          variable=self.var_anim, font=self.f_small, text_color=CLR["text"],
                          progress_color=CLR["accent"],
                          button_color=CLR["card"]).pack(anchor="w")
            ctk.CTkLabel(body, text="低配机器 / 远程桌面觉得卡的时候，把它关掉即可（切页本身从不做动画）",
                         font=self.f_small, text_color=CLR["sub"]).pack(anchor="w", pady=(8, 0))

            bar = ctk.CTkFrame(page, fg_color="transparent")
            bar.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 16))
            ctk.CTkButton(bar, text="保存设置", width=120, height=38, corner_radius=9, font=self.f_btn,
                          fg_color=CLR["accent"], hover_color=CLR["accent_h"], text_color="#FFFFFF",
                          command=self._save_settings_form).pack(side="left")
            ctk.CTkButton(bar, text="恢复默认", width=96, height=38, corner_radius=9, font=self.f_small,
                          fg_color=CLR["ghost"], hover_color=CLR["ghost_h"], text_color=CLR["text"],
                          command=self._reset_settings_form).pack(side="left", padx=(10, 0))
            self.lbl_settings_msg = ctk.CTkLabel(bar, text="", font=self.f_small, text_color=CLR["sub"])
            self.lbl_settings_msg.pack(side="left", padx=(14, 0))
            return page

        def _section(self, parent, title, desc=""):
            card = ctk.CTkFrame(parent, fg_color=CLR["card"], corner_radius=12,
                                border_width=1, border_color=CLR["line"])
            card.pack(fill="x", padx=2, pady=(0, 12))
            ctk.CTkLabel(card, text=title, font=self.f_sec,
                         text_color=CLR["text"]).pack(anchor="w", padx=16, pady=(12, 0))
            if desc:
                ctk.CTkLabel(card, text=desc, font=self.f_small,
                             text_color=CLR["sub"]).pack(anchor="w", padx=16, pady=(2, 0))
            body = ctk.CTkFrame(card, fg_color="transparent")
            body.pack(fill="x", padx=16, pady=(8, 12))
            return body

        def _setting_row(self, body, label):
            row = ctk.CTkFrame(body, fg_color="transparent")
            row.pack(fill="x", pady=5)
            ctk.CTkLabel(row, text=label, width=96, anchor="w", font=self.f_body,
                         text_color=CLR["text"]).pack(side="left")
            return row

        def _entry(self, parent, width=180):
            ent = ctk.CTkEntry(parent, width=width, height=32, font=self.f_body, corner_radius=8,
                               fg_color=CLR["log_bg"], border_color=CLR["line"],
                               text_color=CLR["text"])
            ent.pack(side="left")
            return ent

        def _load_settings_into_form(self):
            self._cfg = load_settings(force=True)
            filters = self._cfg.get("filters") or {}
            for key, widget in self._set_vars.items():
                if key.startswith("filter::"):
                    col = key.split("::", 1)[1]
                    widget.delete(0, "end")
                    widget.insert(0, str(filters.get(col, "")))
                elif key == "sort_keyword":
                    widget.delete(0, "end")
                    widget.insert(0, self._cfg.get("sort_keyword", SORT_COLUMN_KEYWORD))
                elif key == "dedup_keyword":
                    widget.delete(0, "end")
                    widget.insert(0, self._cfg.get("dedup_keyword", DEDUP_COLUMN_KEYWORD))
                elif key == "clipboard_range":
                    widget.delete(0, "end")
                    widget.insert(0, self._cfg.get("clipboard_range") or CLIPBOARD_RANGE)
            try:
                self.var_clip.set(bool(self._cfg.get("clipboard_enabled", True)))
                self.var_popup.set(bool(self._cfg.get("popup_on_done", False)))
                self.var_anim.set(bool(self._cfg.get("animations", True)))
                self.seg_sort.set("降序" if self._cfg.get("sort_desc", True) else "升序")
            except Exception:
                pass

        def _collect_settings(self):
            filters = {}
            for key, widget in self._set_vars.items():
                if key.startswith("filter::"):
                    col = key.split("::", 1)[1]
                    val = widget.get().strip()
                    if val:
                        filters[col] = val
            return {
                "filters": filters or dict(FILTER_COLUMNS),
                "sort_keyword": self._set_vars["sort_keyword"].get().strip() or SORT_COLUMN_KEYWORD,
                "sort_desc": self.seg_sort.get() != "升序",
                "dedup_keyword": self._set_vars["dedup_keyword"].get().strip() or DEDUP_COLUMN_KEYWORD,
                "clipboard_range": self._set_vars["clipboard_range"].get().strip() or CLIPBOARD_RANGE,
                "clipboard_enabled": bool(self.var_clip.get()),
                "popup_on_done": bool(self.var_popup.get()),
                "animations": bool(self.var_anim.get()),
            }

        def _save_settings_form(self):
            data = self._collect_settings()
            self._cfg = save_settings(data)
            self._preview_range = self._cfg.get("clipboard_range") or CLIPBOARD_RANGE
            self.lbl_range.configure(text=self._preview_range)
            if not self._anim_on():        # 关掉动效就把正在跑的补间停掉
                self._stop_anims()
                self.lbl_status.configure(fg_color=CLR["ghost"])
            self.lbl_settings_msg.configure(text="已保存", text_color=OK_CLR)
            self.after(1800, lambda: self.lbl_settings_msg.configure(text=""))

        def _reset_settings_form(self):
            self._cfg = save_settings(json.loads(json.dumps(DEFAULT_SETTINGS)))
            self._load_settings_into_form()
            self._preview_range = self._cfg.get("clipboard_range") or CLIPBOARD_RANGE
            self.lbl_range.configure(text=self._preview_range)
            self.lbl_settings_msg.configure(text="已恢复默认", text_color=WARN_CLR)
            self.after(1800, lambda: self.lbl_settings_msg.configure(text=""))

        # ---------------- 交互：日期 / 主题 ----------------
        def _set_quick_day(self, offset):
            day = (date.today() - timedelta(days=offset)).strftime("%Y%m%d")
            self.entry_day.delete(0, "end")
            self.entry_day.insert(0, day)

        def _open_calendar(self):
            day = re.sub(r"[^0-9]", "", self.entry_day.get()) or self._day_default
            try:
                DatePickerPopup(self, self.entry_day, day, on_pick=self._on_day_picked)
            except Exception as e:
                log(f"[警告] 日历打开失败: {e}")

        def _on_day_picked(self, day):
            self.entry_day.delete(0, "end")
            self.entry_day.insert(0, day)

        def _on_theme_change(self, value):
            mode = "dark" if value == "深色" else "light"
            ctk.set_appearance_mode(mode)
            self._cfg = save_settings({"theme": mode})

            def _repaint():
                # 预览区是 tk.Canvas 自绘的，背景色写死在创建时，换主题必须重建一次
                self._mount_preview_canvas()

            self.after(60, _repaint)

        # ---------------- 日志 ----------------
        def _toggle_log(self):
            self._log_open = not getattr(self, "_log_open", True)
            if self._log_open:
                self.txt.grid()
                self.btn_log.configure(text="运行日志  ▴")
            else:
                self.txt.grid_remove()
                self.btn_log.configure(text="运行日志  ▾")

        def _clear_log(self):
            self.txt.configure(state="normal")
            self.txt.delete("1.0", "end")
            self.txt.configure(state="disabled")

        def _copy_log(self):
            text = self.txt.get("1.0", "end").strip()
            if not text:
                return
            try:
                set_clipboard([[line] for line in text.splitlines()])
                self._flash_status("日志已复制", OK_CLR)
            except Exception as e:
                log(f"[警告] 复制日志失败: {e}")

        def _append_log(self, text):
            self.txt.configure(state="normal")
            self.txt.insert("end", text + "\n")
            self.txt.see("end")
            self.txt.configure(state="disabled")

        # ---------------- 动效（补间引擎，设置里可整体关闭） ----------------
        # 只用 tkinter 的 after 做数值补间，不做贴图/透明，避免这套 GUI 出现掉帧：
        #   · 进度条按真实抓取比例平滑推进（而不是干等一个来回滚的条纹）
        #   · 统计数字滚动到目标值
        #   · 运行中状态胶囊轻轻呼吸
        # 所有补间都按 key 登记，重入时先取消旧的；窗口关闭 / 任务收尾统一停掉。
        ANIM_TICK_MS = 16          # ≈60fps，tkinter 在这个节奏下最稳
        ANIM_STAT_MS = 480         # 统计数字滚动时长
        ANIM_BAR_MS = 260          # 进度条推进时长
        ANIM_BAR_TICK_MS = 24      # 进度条≈40fps：4px 细条看不出差别，省掉 1/3 次重绘
        ANIM_PULSE_MS = 620        # 状态胶囊呼吸节奏

        def _anim_on(self):
            return bool(self._cfg.get("animations", True))

        @staticmethod
        def _ease_out(t):
            """ease-out cubic：开头快、收尾轻，数字滚动手感比线性自然"""
            return 1 - (1 - t) ** 3

        def _cancel_tween(self, key):
            aid = self._tweens.pop(key, None)
            if aid:
                try:
                    self.after_cancel(aid)
                except Exception:
                    pass

        def _tween(self, key, start, end, duration, step, done=None, tick=None):
            """把 start→end 在 duration 毫秒内按 ease-out 喂给 step"""
            self._cancel_tween(key)
            if not self._anim_on() or duration <= 0 or abs(end - start) < 1e-9:
                try:
                    step(end)
                except Exception:
                    pass
                if done:
                    done()
                return
            tick = tick or self.ANIM_TICK_MS
            t0 = time.monotonic()

            def frame():
                t = min(1.0, (time.monotonic() - t0) * 1000.0 / duration)
                try:
                    step(start + (end - start) * self._ease_out(t))
                except Exception:
                    self._tweens.pop(key, None)
                    return
                if t < 1.0:
                    self._tweens[key] = self.after(tick, frame)
                else:
                    self._tweens.pop(key, None)
                    if done:
                        done()

            frame()

        def _set_stat(self, key, value, animate=True):
            """统计块数字：从当前显示值滚到目标值（千位以上带逗号）"""
            lbl = self._stat[key]
            try:
                target = int(value)
            except (TypeError, ValueError):
                lbl.configure(text=str(value))
                return
            raw = re.sub(r"[^0-9]", "", str(lbl.cget("text")))
            current = int(raw) if raw else 0
            if not animate or not self._anim_on() or abs(target - current) < 2:
                lbl.configure(text=f"{target:,}")
                return
            self._tween(f"stat:{key}", current, target, self.ANIM_STAT_MS,
                        lambda v: lbl.configure(text=f"{int(round(v)):,}"))

        # ---- 进度条 ----
        def _progress_begin(self):
            self._cancel_tween("progress")
            self._progress_val = 0.0
            try:
                self.progress.stop()
                self.progress.configure(mode="determinate")
                self.progress.set(0)
            except Exception:
                pass

        def _progress_set(self, value):
            self._progress_val = value
            try:
                self.progress.set(value)
            except Exception:
                pass

        def _progress_to(self, frac):
            """推到 frac（0~1）；从不定态切回来时先停下条纹动画"""
            frac = max(0.0, min(1.0, float(frac)))
            try:
                if self.progress.cget("mode") != "determinate":
                    self.progress.stop()
                    self.progress.configure(mode="determinate")
            except Exception:
                pass
            start = getattr(self, "_progress_val", 0.0)
            if frac < start:            # 回退（重跑）时直接落位，别让条往回爬
                start = frac
            self._tween("progress", start, frac, self.ANIM_BAR_MS, self._progress_set,
                        tick=self.ANIM_BAR_TICK_MS)

        def _progress_busy(self):
            """没有总量可算的阶段（导出 / 处理 / 复制）→ 换成 indeterminate 来回滚"""
            self._cancel_tween("progress")
            try:
                if self.progress.cget("mode") != "indeterminate":
                    self.progress.configure(mode="determinate")
                    self.progress.set(max(getattr(self, "_progress_val", 0.0), 0.06))
                    self.progress.configure(mode="indeterminate")
                    self.progress.start()
            except Exception:
                pass

        # ---- 状态胶囊呼吸 ----
        def _pill_pulse(self, on):
            self._pill_busy = bool(on)
            self._cancel_tween("pill_beat")
            if not on or not self._anim_on():
                try:
                    self.lbl_status.configure(fg_color=CLR["ghost"])
                except Exception:
                    pass
                return

            def beat(i=0):
                if not self._pill_busy:
                    return
                try:
                    self.lbl_status.configure(
                        fg_color=CLR["nav_on"] if i % 2 == 0 else CLR["ghost"])
                except Exception:
                    return
                self._tweens["pill_beat"] = self.after(
                    self.ANIM_PULSE_MS, lambda: beat(i + 1))

            beat(0)

        def _stop_anims(self):
            self._pill_busy = False
            for key in list(self._tweens):
                self._cancel_tween(key)

        # ---------------- 队列 ----------------
        def _drain(self):
            try:
                while True:
                    kind, payload = self._q.get_nowait()
                    if kind == "log":
                        self._append_log(payload)
                    elif kind == "done":
                        self._on_finished(payload)
                    elif kind == "error":
                        self._on_failed(payload[0], payload[1])
                    elif kind == "ask_login":
                        self._show_login(payload)
                    elif kind == "license_ok":
                        self._on_license_ok(payload)
                    elif kind == "license_fail":
                        self._on_license_fail(payload)
                    elif kind == "update":
                        self._show_update_dialog(payload)
                    elif kind == "update_progress":
                        if self._update_dialog is not None:
                            self._update_dialog._on_progress(payload[0], payload[1])
                    elif kind == "update_stage":
                        if self._update_dialog is not None:
                            self._update_dialog._on_stage(payload[0], payload[1])
                    elif kind == "update_ready":
                        if self._update_dialog is not None:
                            self._update_dialog._on_ready(payload)
                    elif kind == "update_error":
                        if self._update_dialog is not None:
                            self._update_dialog._on_error(payload)
                    elif kind == "check_state":
                        self.btn_check.configure(state=payload[0], text=payload[1])
                    elif kind == "progress":
                        self._on_progress(*payload)
                    elif kind == "hist_rows":
                        lbl = self._hist_rows.get(payload[0])
                        if lbl is not None:
                            lbl.configure(text=str(payload[1]))
            except queue.Empty:
                pass
            self.after(80, self._drain)

        # ---------------- 运行 ----------------
        def _get_day(self):
            return re.sub(r"[^0-9]", "", self.entry_day.get())

        def _on_start(self):
            if self._running:
                return
            day = self._get_day()
            try:
                datetime.strptime(day, "%Y%m%d")
            except ValueError:
                self.entry_day.configure(border_color=ERR_CLR)
                self.after(1500, lambda: self.entry_day.configure(border_color=CLR["line"]))
                self._append_log(f"{datetime.now():%H:%M:%S}  [提示] 日期格式不正确，应为 20261119")
                self._set_status("日期有误", ERR_CLR)
                return
            if not self._license_ok:
                self._pending_day = day
                return

            self._running = True
            self._last_day = day
            self.btn_start.configure(state="disabled", text="核查中…")
            self._progress_begin()
            self._pill_pulse(True)
            self._set_status("运行中", RUN_CLR)
            self._reset_stages()
            self._append_log("")
            threading.Thread(target=self._worker, args=(day,), daemon=True).start()

        def _worker(self, day):
            def progress(stage, done, total):
                self._q.put(("progress", (stage, done, total)))
            try:
                res = run_pipeline(day, progress=progress)
            except PipelineError as e:
                self._q.put(("error", (e.stage, str(e))))
                return
            except Exception as e:
                report_exception(e, stage="unknown", day=str(day))
                self._q.put(("error", ("unknown", str(e))))
                return
            self._q.put(("done", res))

        def _on_progress(self, stage, done, total):
            self._paint_stage(stage, done, total)
            if stage == "fetch":
                if total:
                    # 有总量才能算真实百分比：进度条按 抓到的条数/总条数 平滑推进
                    self.lbl_stage_note.configure(text=f"{done}/{total} 条")
                    self._progress_to(done / float(total))
                else:
                    self.lbl_stage_note.configure(text=f"已抓取 {done} 条" if done else "")
                    self._progress_busy()
            else:
                # 导出 / 处理 / 复制：没有可比的总量，换成来回滚的条纹
                self.lbl_stage_note.configure(text="")
                self._progress_busy()

        def _reset_stages(self):
            for key, lbl in self._stage_lbl.items():
                lbl.configure(text=self.STAGE_TEXT[key], text_color=CLR["sub"])
            self.lbl_stage_note.configure(text="")

        def _paint_stage(self, current, done=0, total=0):
            """阶段指示只改颜色：文字宽度不变，避免每来一批数据就把整行挤得左右跳"""
            try:
                idx_cur = self.STAGE_ORDER.index(current)
            except ValueError:
                idx_cur = -1
            for i, key in enumerate(self.STAGE_ORDER):
                lbl = self._stage_lbl[key]
                text = self.STAGE_TEXT[key]
                if i < idx_cur:
                    lbl.configure(text=text, text_color=OK_CLR)
                elif i == idx_cur:
                    lbl.configure(text=text, text_color=CLR["accent"])
                else:
                    lbl.configure(text=text, text_color=CLR["sub"])

        def _set_status(self, text, color):
            # 序号自增：任何一次状态变更都会让之前挂着的「自动恢复就绪」失效，
            # 否则「已复制」的 2 秒复位可能把后来的「运行中」「处理失败」冲掉。
            self._status_seq = getattr(self, "_status_seq", 0) + 1
            try:
                self.lbl_status.configure(text=f"●  {text}", text_color=color)
            except Exception:
                pass
            return self._status_seq

        def _flash_status(self, text, color):
            seq = self._set_status(text, color)

            def _restore():
                if getattr(self, "_status_seq", 0) == seq:
                    self._set_status("就绪", CLR["sub"])

            self.after(2200, _restore)

        def _on_finished(self, res):
            self._running = False
            self.btn_start.configure(state="normal", text="开始核查")
            self._reset_progress(True)
            stat = res.get("stat") or {}
            if res.get("status") == "ok":
                self._set_stat("fetch", res.get("count", 0))
                self._set_stat("filter", stat.get("rows_after_filter", 0))
                self._set_stat("dedup", stat.get("rows_after_dedup", 0))
                self._last_out = stat.get("out_path") or ""
                clip = stat.get("clip") or {}
                if clip.get("ok"):
                    self.clip_lbl.configure(
                        text=f"已复制 {clip.get('range')} 到剪贴板（{clip.get('rows')} 行 × {clip.get('cols')} 列）")
                    self.clip_bar.grid()
                    self._set_status("已完成", OK_CLR)
                else:
                    self.clip_bar.grid_remove()
                    self._set_status("已完成", OK_CLR)
                self._render_preview(self._read_preview(self._last_out),
                                     (clip.get("range") if clip.get("ok") else self._preview_range))
                self._refresh_recent()
                self._append_log(f"{datetime.now():%H:%M:%S}  完成：{stat.get('rows_after_dedup', 0)} 行")
            else:
                self._set_stat("fetch", 0, animate=False)
                self._stat["filter"].configure(text="—")
                self._stat["dedup"].configure(text="—")
                self.clip_bar.grid_remove()
                self._set_status("无数据", WARN_CLR)
                self._render_preview([], self._preview_range)
            if self._cfg.get("popup_on_done", False):
                try:
                    ResultWindow(self, res)
                except Exception as e:
                    log(f"[警告] 结果窗口打开失败: {e}")

        def _read_preview(self, path):
            if not path or not os.path.isfile(path):
                return []
            try:
                return read_xlsx_range(path, self._preview_range)
            except Exception as e:
                log(f"[警告] 预览读取失败: {e}")
                return []

        def _on_failed(self, stage, msg):
            self._running = False
            self.btn_start.configure(state="normal", text="开始核查")
            self._reset_progress(False)
            stage_text = {"fetch": "抓取失败", "csv": "导出失败", "process": "处理失败"}.get(stage, "运行失败")
            self._set_status(stage_text, ERR_CLR)
            self._paint_stage("", 0, 0)
            try:
                import tkinter.messagebox as mb
                mb.showerror("运行失败", f"{stage_text}\n\n{msg}")
            except Exception:
                pass

        def _reset_progress(self, done=True):
            """收尾：停掉呼吸/条纹；成功时把条平滑推到满格，失败时停在当前进度不动"""
            self._pill_pulse(False)
            if done:
                self._progress_to(1.0)
            else:
                self._cancel_tween("progress")
                try:
                    self.progress.stop()
                except Exception:
                    pass

        # ---------------- 复制 / 目录 ----------------
        def _copy_again(self):
            if not self._last_out:
                return
            clip = copy_xlsx_range_to_clipboard(self._last_out)
            if clip["ok"]:
                self.clip_lbl.configure(
                    text=f"已复制 {clip['range']} 到剪贴板（{clip['rows']} 行 × {clip['cols']} 列）")
                self._flash_status("已复制", OK_CLR)
            else:
                self._flash_status("复制失败", ERR_CLR)
                log(f"[警告] 复制失败：{clip['error']}")

        def _copy_for(self, path):
            clip = copy_xlsx_range_to_clipboard(path)
            if clip["ok"]:
                self._flash_status("已复制", OK_CLR)
                log(f"[剪贴板] 已复制 {os.path.basename(path)} 的 {clip['range']}")
            else:
                self._flash_status("复制失败", ERR_CLR)
                log(f"[警告] 复制失败：{clip['error']}")

        def _open_output_folder(self):
            self._open_path(output_folder())

        def _open_path(self, path):
            try:
                open_folder(path)
            except Exception as e:
                log(f"[警告] 打开目录失败: {e}")

        # ---------------- 最近任务 ----------------
        def _refresh_recent(self):
            for w in self.recent_box.winfo_children():
                w.destroy()
            items = scan_outputs(limit=3)
            if not items:
                ctk.CTkLabel(self.recent_box, text="暂无产物", font=self.f_small,
                             text_color=CLR["sub"]).pack(anchor="w", padx=6)
                return
            for it in items:
                row = ctk.CTkFrame(self.recent_box, fg_color="transparent", height=24)
                row.pack(fill="x")
                ctk.CTkLabel(row, text="●", font=self.f_small, text_color=OK_CLR,
                             width=12).pack(side="left", padx=(6, 0))
                text = f"{it['day'][4:6]}-{it['day'][6:]}  " + (
                    f"{it['rows'] - 1} 行" if it["rows"] else "已产出")
                ctk.CTkLabel(row, text=text, font=self.f_small,
                             text_color=CLR["text"]).pack(side="left")

        # ---------------- 授权（后台静默，界面不展示） ----------------
        def _reveal_for_license(self):
            if getattr(self, "_ui_hidden", False):
                self._ui_hidden = False
                try:
                    self.deiconify()
                    self.lift()
                except Exception:
                    pass

        def _start_license_check(self):
            if not LICENSE_ENABLED:
                self._license_ok = True
                self._license_pending = False
                _license_trace("已按配置跳过授权校验")
                self._reveal_for_license()
                self._schedule_update_check()
                return
            self._license_ok = False
            self._license_pending = True
            threading.Thread(target=self._license_worker, daemon=True).start()

        def _license_worker(self):
            denied = {}
            try:
                ok, payload, guard = verify_license(
                    on_denied=lambda c, m: denied.update(code=c, message=m))
            except Exception as e:
                ok, payload, guard = False, ("UNKNOWN", f"授权校验异常: {e}"), None
            if not ok and denied:
                c, m = denied.get("code"), denied.get("message")
                payload = (c, license_hint(c, m))
            self._license_guard = guard
            self._q.put(("license_ok" if ok else "license_fail", payload))

        def _on_license_ok(self, info):
            self._license_ok = True
            self._license_pending = False
            self._license_info = info
            self._reveal_for_license()
            who = (info.customer if info else None) or "未署名"
            if info is not None and info.hard_expires_at:
                tail = f"剩余 {info.days_left()} 天"
            else:
                tail = "永久有效"
            if info is not None and info.offline:
                tail += "（本次为离线判定）"
            _license_trace(f"校验通过：{who} · {tail}")
            self._start_license_watchdog()
            self._schedule_update_check()
            if self._pending_day:
                self._pending_day = None
                self.after(10, self._on_start)

        def _on_license_fail(self, payload):
            code, msg = payload
            self._license_ok = False
            self._license_pending = False
            self._pending_day = None
            _, exit_code = license_fail_behavior(self._license_guard)
            _license_trace(f"校验未通过: [{code}] {msg}（静默退出，退出码 {exit_code}）")
            self.after(200, lambda: _exit_now(exit_code))

        def _start_license_watchdog(self):
            guard = self._license_guard
            if guard is None:
                return
            try:
                guard.cfg.on_denied = lambda c, m: self._q.put(
                    ("license_fail", (c, license_hint(c, m))))
                guard.start_watchdog()
            except Exception:
                pass

        # ---------------- 在线更新 ----------------
        def _schedule_update_check(self):
            if not UPDATE_AUTO_CHECK or os.environ.get("AIREPLAY_NO_UPDATE"):
                return
            if is_dev_build():
                _license_trace("开发构建（版本号未注入），跳过自动更新检查")
                return
            self.after(1200, lambda: self._check_update(manual=False))

        def _check_update(self, manual=False):
            try:
                self.btn_check.configure(state="disabled", text="检测中…")
            except Exception:
                pass
            if manual:
                self._append_log(f"{datetime.now():%H:%M:%S}  [更新] 正在检查新版本…")
            threading.Thread(target=self._update_worker, args=(manual,), daemon=True).start()

        def _update_worker(self, manual):
            try:
                status, payload = check_for_update()
            except Exception as e:
                status, payload = "error", f"{type(e).__name__}: {e}"
            ts = f"{datetime.now():%H:%M:%S}"
            if status == "update":
                self._q.put(("update", payload))
            elif manual:
                if status == "latest":
                    self._q.put(("log", f"{ts}  [更新] 已是最新版本 v{payload.get('version', '')}"))
                else:
                    self._q.put(("log", f"{ts}  [更新] 检测失败：{payload}"))
            self._q.put(("check_state", ("normal", "检查更新")))

        def _show_update_dialog(self, info):
            self._append_log(
                f"{datetime.now():%H:%M:%S}  [更新] 发现新版本 v{info['version']}"
                f"（当前 v{current_app_version()}），可一键升级")
            try:
                if self._update_dialog is not None:
                    try:
                        self._update_dialog.destroy()
                    except Exception:
                        pass
                self._update_dialog = UpdateDialog(self, info)
            except Exception as e:
                log(f"[警告] 更新窗口打开失败: {e}")

        def _quit_for_update(self):
            try:
                self.destroy()
            except Exception:
                pass
            _exit_now(0)

        # ---------------- 登录 / 预热 / 关闭 ----------------
        def _ask_login_threadsafe(self, tip="", error="", default_account=""):
            req = {"tip": tip, "error": error, "account": default_account,
                   "res": None, "event": threading.Event()}
            self._q.put(("ask_login", req))
            req["event"].wait()
            return req["res"]

        def _show_login(self, req):
            try:
                req["res"] = LoginDialog(
                    self, tip=req.get("tip", ""), error=req.get("error", ""),
                    default_account=req.get("account", ""),
                ).wait()
            except Exception as e:
                log(f"[警告] 登录窗口打开失败: {e}")
                req["res"] = None
            finally:
                req["event"].set()

        def _preheat(self):
            # 空闲时先把 pandas 拉进来（首次 import 要几百毫秒），
            # 免得点「开始」那一刻才付这笔钱
            try:
                import importlib
                importlib.import_module("pandas")
            except Exception:
                pass

        def _on_close(self):
            if self._running:
                try:
                    import tkinter.messagebox as mb
                    if not mb.askyesno("确认退出", "任务正在运行中，确定要退出吗？"):
                        return
                except Exception:
                    pass
            self._stop_anims()          # 先摘掉所有补间定时器，否则销毁后回调会报错
            try:
                self.progress.stop()    # 停掉进度条自己的循环动画
            except Exception:
                pass
            self.destroy()

    def launch_gui(selfcheck=False):
        """启动图形界面；selfcheck=True 时短暂启动后自动退出（用于打包自检）"""
        # 首帧之前就把外观定成用户上次选的（默认浅色），否则会先闪一下深色再切过来
        ctk.set_appearance_mode("dark" if load_settings().get("theme") == "dark" else "light")
        ctk.set_default_color_theme("blue")
        app = ReplayApp(selfcheck=selfcheck)
        if selfcheck:
            app.after(1800, lambda: (app.destroy(), _console("SELFCHECK-OK")))
        app.mainloop()
        return 0


# ==================================================================
#                              入口
# ==================================================================
def _parse_args(argv):
    """返回 (day, selfcheck, action)；
    day 为 YYYYMMDD 或 None，action ∈ {None, "check_update", "version"}"""
    day, selfcheck, action, rest = None, False, None, []
    i = 0
    while i < len(argv):
        a = argv[i]
        if a in ("--selfcheck", "-s"):
            selfcheck = True
        elif a in ("--check-update",):
            action = "check_update"
        elif a in ("--version", "-V"):
            action = "version"
        elif a in ("--cli", "--date", "-d"):
            i += 1
            if i < len(argv):
                day = re.sub(r"[^0-9]", "", argv[i])
        elif a.startswith("--date="):
            day = re.sub(r"[^0-9]", "", a.split("=", 1)[1])
        elif a in ("--gui", "-g"):
            pass
        else:
            rest.append(a)
        i += 1
    if day is None and rest:
        day = re.sub(r"[^0-9]", "", rest[0])
    return day, selfcheck, action


def main():
    # 自检开关: 只执行一次「打开输出目录」动作后退出（用于打包产物离线验证，不影响正常使用）
    if "--open-out" in sys.argv[1:]:
        try:
            open_folder(os.path.join(script_dir(), "replay核查"))
            return 0
        except Exception:
            return 1

    day_arg, selfcheck, action = _parse_args(sys.argv[1:])

    # --version：打印当前版本号后退出（便于 CI 校验注入结果）
    if action == "version":
        _console(current_app_version())
        return 0

    # 注意：以前这里有个 --license 子命令，会在控制台打印授权状态。
    # 现在不允许出现任何能看出软件带授权的字样，该子命令已删除。
    # 排障需要看授权结果时，设环境变量 AIREPLAY_LICENSE_TRACE=<文件路径>。

    # --check-update：只检测更新并打印结果，不启动界面
    if action == "check_update":
        status, payload = check_for_update()
        if status == "update":
            _console(f"NEW={payload['version']}  FILE={payload['asset_name']}")
        elif status == "latest":
            _console(f"LATEST  current={current_app_version()}")
        else:
            _console(f"ERROR  {payload}")
        return 0 if status in ("update", "latest") else 1

    # 上次在线更新的收尾：清理旧版本文件（后台进行，不影响启动）
    if not selfcheck:
        threading.Thread(target=cleanup_after_update, daemon=True).start()

    # 1. 命令行静默模式（传日期）
    if day_arg:
        rc = enforce_license_cli()     # 静默模式同样要过授权，否则等于可以绕开校验
        if rc:
            return rc
        try:
            datetime.strptime(day_arg, "%Y%m%d")
        except ValueError:
            _console(f"日期参数格式错误: {day_arg}，请使用 20260924 格式。")
            return 2
        try:
            res = run_pipeline(day_arg)
        except PipelineError:
            return 1
        except Exception as e:
            report_exception(e, stage="startup")
            _console(f"[错误] {e}")
            return 1
        return 0 if res.get("status") in ("ok", "empty") else 1

    # 2. 图形界面模式（有 customtkinter）
    if ctk is not None:
        try:
            return launch_gui(selfcheck=selfcheck)
        except Exception as e:
            report_exception(e, stage="gui")
            try:
                with open(os.path.join(script_dir(), "error.log"), "a", encoding="utf-8") as f:
                    f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} 界面异常: {e}\n")
            except Exception:
                pass
            return 1

    # 3. 降级：原生弹窗交互
    rc = enforce_license_cli()
    if rc:
        return rc
    default_day = (date.today() - timedelta(days=1)).strftime("%Y%m%d")
    day = ask_date(default_day)
    if not day:
        log("已取消。")
        return 1
    try:
        res = run_pipeline(day)
    except PipelineError as e:
        popup("运行失败", str(e), kind="error")
        return 1
    if res.get("status") == "empty":
        popup("无数据", f"{res['start_at']} ~ {res['end_at']} 没有「已核查」记录。")
        return 0
    stat = res["stat"]
    clip = stat.get("clip") or {}
    clip_line = (f"剪贴板: 已复制 {clip.get('range')}"
                 f"（{clip.get('rows')} 行 × {clip.get('cols')} 列）\n"
                 if clip.get("ok") else
                 f"剪贴板: 未复制{('（' + clip['error'] + '）') if clip.get('error') else ''}\n")
    popup(
        "文档处理完成",
        f"CSV 记录: {stat['rows_before']} 条\n"
        f"筛选: {stat['rows_before']} → {stat['rows_after_filter']} 行\n"
        f"去重: {stat['rows_before_dedup']} → {stat['rows_after_dedup']} 行（按「{DEDUP_COLUMN_KEYWORD}」）\n"
        f"排序: 按「分数」降序\n"
        f"{clip_line}"
        f"─────────────────\n"
        f"输出文件:\n{stat['out_path']}",
    )
    return 0


if __name__ == "__main__":
    _rc = 0
    try:
        _rc = main()
    except SystemExit as _e:
        _rc = _e.code
    except Exception as e:
        report_exception(e, stage="startup")
        _console(f"[错误] {e}")
        _rc = 1
    # 收尾即退出：关窗口 / 任务跑完后进程立刻消失
    _exit_now(_rc)
