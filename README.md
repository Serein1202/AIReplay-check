<p align="center"><img src="assets/app.png" width="96" alt=""></p>

# AIReplay-check

Replay 核查一体化工具：一次运行完成「后台抓取 → CSV 导出 → 筛选 / 排序 / 去重 → xlsx 输出」，带图形操作界面。

生成 xlsx 后**不打开文件**，自动把报表的 **C1:L11** 区域复制到剪贴板（文本 + HTML 表格两种格式：粘到 Excel 按单元格铺开，粘到微信 / 飞书 / Word 直接成表格）。数据不足 11 行时自动裁掉尾部空行空列。复制区域与开关都在**设置页**里改（立即生效，无需重新发版），工作台的「复制内容预览」会实时显示将被复制的表格。

[![Build EXE & Publish Release](https://github.com/Serein1202/AIReplay-check/actions/workflows/build-release.yml/badge.svg)](https://github.com/Serein1202/AIReplay-check/actions/workflows/build-release.yml)

## 界面

左侧固定导航 + 右侧内容区，三个页面：

| 页面 | 作用 |
| --- | --- |
| **工作台** | 选日期（快捷「昨天 / 今天 / 前天」或点开自绘月历）→ 开始核查；四段阶段指示（抓取 → 导出 → 处理 → 复制）+ 细进度条；三块统计（抓取条数 / 筛选后 / 去重后）；成功后就地出现「已复制 C1:L11」提示条与「再次复制」；下方是复制内容预览表（可横竖滚动）与可收起的运行日志 |
| **历史记录** | 扫输出目录里已有的 `replay核查-<日期>.xlsx`，显示行数 / 大小 / 修改时间，可对任意一份再次复制或打开所在目录（行数后台异步回填，不卡界面） |
| **设置** | 筛选条件、排序列与升降序、去重列、剪贴板区域与开关、完成后是否弹结果窗、界面动效开关；改完点「保存设置」即时生效，「恢复默认」一键回到初始配置 |

左侧底部固定有深/浅色主题切换、检查更新入口与当前版本号；**默认浅色**，切成深色后会记住。窗口默认 1180×800，最小 1080×720。

## 下载使用

到 **[Releases](https://github.com/Serein1202/AIReplay-check/releases/latest)** 页面下载 `AIReplay-<版本号>.exe`（例如 `AIReplay-v1.0.1.exe`），双击运行即可，**无需安装 Python**。

首次运行需要在弹窗中登录以获取 Token。Token 会缓存在 exe 同目录的 `.amis_token.json`，**内含凭据，请勿外传或提交到仓库**（已在 `.gitignore` 中忽略）。

## 软件授权

程序内置授权校验（License Guard）。核心原则是相关参数完全跟随后台设置，
客户端不自己揣测策略，后台改完即生效。

- 授权码不打进仓库：仓库里 `LICENSE_KEY` 是空的，正式构建时由 CI 从仓库 Secret
  `APP_LICENSE_KEY` 注入后再打包 —— 这样公开仓库和 git 历史里都查不到授权码，
  换码也只需改 Secret 重跑构建，不必改代码。运行时的覆盖优先级：
  环境变量 `LICENSE_KEY` > 程序同目录的 `license.key` > 脚本内置值
- 每次启动都会联网校验（`LICENSE_ONLINE_CHECK_SEC = 0`），因此后台把授权码改成
  过期 / 吊销 / 暂停，客户端下次启动立刻被拦下（此前版本会被本地缓存令牌「离线放行」挡掉）
- 以下参数由后台下发，客户端原样执行：

  | 后台设置 | 客户端行为 |
  |---|---|
  | 到期时的行为（直接闪退 / 弹提示后退出） | 完全照做：`hard` 静默退出、`message` 弹一次提示再退出 |
  | 心跳间隔 | 运行期复查间隔就取这个值（本地 `LICENSE_HEARTBEAT_CAP_SEC = 0` 表示不封顶） |
  | 断网容忍天数 | 离线放行的窗口；设为 `0` 即完全不允许离线放行 |
  | 到期后宽限天数 / 设备数上限 / 令牌有效期 | 服务端计算，客户端只验签 |
  | 签名算法 / app_id | 与脚本里的 `LICENSE_SIGN_ALG` / `LICENSE_APP_ID` 一致 |

- 只有真的连不上授权服务器时，才按后台的「断网容忍天数」离线放行
- 服务端明确判失效（过期 / 吊销 / 暂停 / 授权码不存在）时会同时清掉本地缓存令牌，
  避免「断网后靠旧令牌继续跑」
- 界面不显示任何授权信息，也没有查看授权状态的命令：失败原因只走 stdout，而 exe 是
  `--noconsole` 构建，等同于不可见。排障时设环境变量
  `AIREPLAY_LICENSE_TRACE=<日志文件路径>` 才会把校验过程写盘

> 注意：后台表单里的「断网容忍天数」若设为 `0`，程序就必须每次启动都能连上授权服务器。
> 本机到授权服务的链路曾出现整段时间不可达（超时），此时程序会因 `NETWORK` 而无法启动。

> `license.key` 若存在则含凭据，请勿提交到仓库（已在 `.gitignore` 中忽略）。

## 自动更新检测

程序打开时会自动检测 GitHub 最新 Release：

- 发现更高版本 → 弹窗显示新旧版本、安装包大小与更新说明，「立即更新」即自动下载并重启到新版本
- 新版本下载为 `AIReplay-<版本号>.exe` 放在原程序同目录，重启后自动清理旧版本文件，目录里不会堆积历史 exe
- 下载采用「先写 `.part` 再改名」并校验字节数，中断或文件不完整会拒绝替换
- 无法访问 `api.github.com`（网络受限等）时静默跳过，完全不影响正常使用

界面左下角显示当前版本号，旁边的「检查更新」可手动触发检测（会在运行日志里输出检测结果）。

关闭自动检测：设置环境变量 `AIREPLAY_NO_UPDATE=1`，或把脚本里的 `UPDATE_AUTO_CHECK` 改为 `False`。

## 本地配置（设置页改的就是它）

设置页保存的筛选 / 排序 / 去重 / 剪贴板 / 主题等参数，写在 exe（或脚本）同目录的
`.aireplay_config.json`，已在 `.gitignore` 中忽略：

- 文件不存在 → 全部用代码里的默认值，行为与老版本完全一致
- 文件损坏或字段缺失 → 缺什么补什么，不会整份覆盖
- 删掉这个文件 = 恢复出厂设置（设置页的「恢复默认」效果相同）

| 字段 | 说明 |
| --- | --- |
| `filters` | 筛选条件，`{"列名关键字": "期望值"}`，需全部满足 |
| `sort_keyword` / `sort_desc` | 按哪列排序；分数列按数值比较 |
| `dedup_keyword` | 去重列关键字，同一角色只保留排序后的第一行 |
| `clipboard_enabled` / `clipboard_range` | 输出后是否自动复制、复制哪个区域（Excel 记法，如 `C1:L11`） |
| `popup_on_done` | 完成后是否弹结果窗口（默认关，避免打断粘贴） |
| `animations` | 界面动效（进度条平滑推进 / 统计数字滚动 / 运行中状态呼吸），关掉则所有状态切换瞬时完成；切页始终是瞬时切换，不受此项影响 |
| `theme` | `dark` / `light` |
| `history_limit` | 历史页最多列出多少份产物 |

## 自动构建与发布

仓库已配置 GitHub Actions：推送到 `main` 分支就会自动打包成 exe 并发布到 Releases，无需本地装环境。

| 操作 | 结果 |
| --- | --- |
| 向 `main` 分支推送提交 | 自动构建，版本号取历史最高 tag 递增一位（patch 只跑 0~9，到 9 进 minor：v1.0.8→v1.0.9→v1.1.0），自动打 tag 并发布 |
| 在 Actions 页面点 **Run workflow** | 留空 = 重发最新已有 tag（不 +1）；填某个已存在的 tag = 重发该版本 |

每次发布会生成一个可执行文件 `AIReplay-<版本号>.exe`（例如 `AIReplay-v1.0.3.exe`）。

需要在仓库 **Settings → Secrets and variables → Actions** 里配置一个 Secret：

| Secret | 说明 |
| --- | --- |
| `APP_LICENSE_KEY` | 内置授权码（形如 `XXXX-XXXX-XXXX-XXXX`）。打包前注入脚本，注入失败会直接让构建失败，不会发出「没有授权码」的 exe |

流程配置见 [`.github/workflows/build-release.yml`](.github/workflows/build-release.yml)：

```
检出代码 → 安装 Python 3.12 → 安装依赖 → 计算版本号 → 递增保护 → 自动打 tag
        → 注入版本号到脚本 → 注入授权码到脚本 → PyInstaller 打包 → 校验产物 → 挂载到 Release
```

> 打包前会把计算出的版本号写入脚本的 `APP_VERSION`，所以每个 exe 都确切知道自己是哪个版本 —— 这是自动更新判断「是否需要升级」的依据。
> 注入只发生在 CI 的运行目录里，不会提交回仓库。

## 本地开发

> 授权码不在仓库里，本地运行需先提供：设置环境变量 `LICENSE_KEY`，或在本目录放一个
> `license.key`（内容即授权码，已在 `.gitignore` 中忽略）。**没有授权码时程序会静默退出
> （无窗口、无提示、退出码 1）** —— 这是刻意设计，不是崩溃。排障可临时设置
> `AIREPLAY_LICENSE_TRACE=<日志文件路径>` 查看具体原因。
>
> 另外需要带 `tkinter` 的 Python：`python -c "import tkinter"` 报 `ModuleNotFoundError`
> 的话（部分精简发行版 / 内嵌 Python 常见）界面起不来，请换官方 CPython 安装包
> （如 `py -3.14 AIReplay-check.py`）。

```bash
pip install -r requirements.txt

export LICENSE_KEY=XXXX-XXXX-XXXX-XXXX       # 或写一份 license.key
python AIReplay-check.py                     # 打开图形界面
python AIReplay-check.py 20260924            # 静默模式（可用于定时任务）
python AIReplay-check.py --selfcheck         # 界面自检
python AIReplay-check.py --check-update      # 只检测更新并打印结果
python AIReplay-check.py --version           # 打印当前版本号
```

## 本地打包 exe

```bash
pip install pyinstaller
export LICENSE_KEY=XXXX-XXXX-XXXX-XXXX       # 本地打包同样要提供授权码
pyinstaller --noconsole --onefile --clean \
  --icon "assets/app.ico" --add-data "assets/app.ico;assets" \
  --collect-all customtkinter --name "AIReplay" AIReplay-check.py
```

产物位于 `dist/AIReplay.exe`。

> 本地打包出来的 exe 会带上当时环境里的授权码；正式发布请走 CI（由 Secret 注入），
> 避免把授权码写进本地脚本或忘记注入而打出打不开的包。

## 目录结构

```
AIReplay-check.py                     # 主程序（图形界面 + 静默模式）
license_guard.py                      # 授权校验 SDK（随程序一起分发）
requirements.txt                      # 运行 / 打包依赖
.github/workflows/build-release.yml   # 自动构建并发布 Release
assets/app.ico                        # 应用图标（exe 图标 + 窗口图标）
assets/app.png                        # 图标的 512px 预览图
tools/make_icon.py                    # 图标生成脚本（改配色 / 换造型后重新跑一次，同时重出上面两个文件）
.aireplay_config.json                 # 运行期生成的本机配置（不入库）
```
