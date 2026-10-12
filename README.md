# AIReplay-check

Replay 核查一体化工具：一次运行完成「后台抓取 → CSV 导出 → 筛选 / 排序 / 去重 → xlsx 输出」，带图形操作界面。

[![Build EXE & Publish Release](https://github.com/Serein1202/AIReplay-check/actions/workflows/build-release.yml/badge.svg)](https://github.com/Serein1202/AIReplay-check/actions/workflows/build-release.yml)

## 下载使用

到 **[Releases](https://github.com/Serein1202/AIReplay-check/releases/latest)** 页面下载 `AIReplay-<版本号>.exe`（例如 `AIReplay-v1.0.1.exe`），双击运行即可，**无需安装 Python**。

首次运行需要在弹窗中登录以获取 Token。Token 会缓存在 exe 同目录的 `.amis_token.json`，**内含凭据，请勿外传或提交到仓库**（已在 `.gitignore` 中忽略）。

## 软件授权

程序内置授权校验（License Guard）。**核心原则：相关参数完全跟随后台设置**，
客户端不自己揣测策略，后台改完即生效。

- **授权码不打进仓库**：仓库里 `LICENSE_KEY` 是空的，正式构建时由 CI 从仓库 Secret
  `APP_LICENSE_KEY` 注入后再打包 —— 这样公开仓库和 git 历史里都查不到授权码，
  换码也只需改 Secret 重跑构建，不必改代码。运行时的覆盖优先级：
  环境变量 `LICENSE_KEY` > 程序同目录的 `license.key` > 脚本内置值
- **每次启动都会联网校验**（`LICENSE_ONLINE_CHECK_SEC = 0`），因此后台把授权码改成
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
- 界面**不显示任何授权信息**：失败原因只走 stdout，而 exe 是 `--noconsole` 构建，等同于不可见
- 命令行检查状态：`AIReplay-check.py --license`（失败时返回非 0 退出码）

> ⚠️ 后台表单里的「断网容忍天数」若设为 `0`，程序就**必须每次启动都能连上授权服务器**。
> 本机到授权服务的链路曾出现整段时间不可达（超时），此时程序会因 `NETWORK` 而无法启动。

> `license.key` 若存在则含凭据，请勿提交到仓库（已在 `.gitignore` 中忽略）。

## 自动更新检测

程序**打开时会自动检测 GitHub 最新 Release**：

- 发现更高版本 → 弹窗显示新旧版本、安装包大小与更新说明，「立即更新」即自动下载并重启到新版本
- 新版本下载为 `AIReplay-<版本号>.exe` 放在原程序同目录，重启后自动清理旧版本文件，目录里不会堆积历史 exe
- 下载采用「先写 `.part` 再改名」并校验字节数，中断或文件不完整会拒绝替换
- 无法访问 `api.github.com`（网络受限等）时静默跳过，完全不影响正常使用

界面右下角显示当前版本号，主按钮旁的「检查更新」可手动触发检测（会输出检测结果）。

关闭自动检测：设置环境变量 `AIREPLAY_NO_UPDATE=1`，或把脚本里的 `UPDATE_AUTO_CHECK` 改为 `False`。

## 自动构建与发布

仓库已配置 GitHub Actions：**只要推送 `.py` 文件的改动，就会自动打包成 exe 并发布到 Releases**，无需本地装环境。

| 操作 | 结果 |
| --- | --- |
| 向 `main` 分支推送提交 | 自动构建，版本号取历史最高 tag 的 patch +1（如 v1.0.2 → v1.0.3），自动打 tag 并发布 |
| 在 Actions 页面点 **Run workflow** | 留空 = 重发最新已有 tag（不 +1）；填某个已存在的 tag = 重发该版本 |

每次发布会生成一个可执行文件 `AIReplay-<版本号>.exe`（例如 `AIReplay-v1.0.3.exe`）。

⚠️ 需要在仓库 **Settings → Secrets and variables → Actions** 里配置一个 Secret：

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
> `license.key`（内容即授权码，已在 `.gitignore` 中忽略）。

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
pyinstaller --noconsole --onefile --clean --collect-all customtkinter --name "AIReplay" AIReplay-check.py
```

产物位于 `dist/AIReplay.exe`。

> 本地打包出来的 exe 会带上**当时环境里的**授权码；正式发布请走 CI（由 Secret 注入），
> 避免把授权码写进本地脚本或忘记注入而打出打不开的包。

## 目录结构

```
AIReplay-check.py                     # 主程序（图形界面 + 静默模式）
license_guard.py                      # 授权校验 SDK（随程序一起分发）
requirements.txt                      # 运行 / 打包依赖
.github/workflows/build-release.yml   # 自动构建并发布 Release
```
