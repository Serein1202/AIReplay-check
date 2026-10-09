# AIReplay-check

Replay 核查一体化工具：一次运行完成「后台抓取 → CSV 导出 → 筛选 / 排序 / 去重 → xlsx 输出」，带图形操作界面。

[![Build EXE & Publish Release](https://github.com/Serein1202/AIReplay-check/actions/workflows/build-release.yml/badge.svg)](https://github.com/Serein1202/AIReplay-check/actions/workflows/build-release.yml)

## 下载使用

到 **[Releases](https://github.com/Serein1202/AIReplay-check/releases/latest)** 页面下载 `AIReplay-<版本号>.exe`（例如 `AIReplay-v1.0.1.exe`），双击运行即可，**无需安装 Python**。

首次运行需要在弹窗中登录以获取 Token。Token 会缓存在 exe 同目录的 `.amis_token.json`，**内含凭据，请勿外传或提交到仓库**（已在 `.gitignore` 中忽略）。

## 软件授权

程序内置授权校验（License Guard）。**核心原则：相关参数完全跟随后台设置**，
客户端不自己揣测策略，后台改完即生效。

- **授权码已内置在程序里**，开箱即用，无需额外配置文件
  （按客户分发要不同授权码时，覆盖优先级：环境变量 `LICENSE_KEY` >
  程序同目录的 `license.key` > 内置授权码）
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
| 向 `main` 分支推送含 `.py` 的改动 | 自动构建，版本号自动递增为 `v1.0.<构建号>`，生成 Release |
| `git tag v1.2.0 && git push origin v1.2.0` | 以 `v1.2.0` 作为版本号构建并发布 |
| 在 Actions 页面点 **Run workflow** | 手动触发，可自定义版本号（留空则自动生成） |

每次发布会生成一个可执行文件 `AIReplay-<版本号>.exe`（例如 `AIReplay-v1.0.1.exe`）。

流程配置见 [`.github/workflows/build-release.yml`](.github/workflows/build-release.yml)：

```
检出代码 → 安装 Python 3.12 → 安装依赖 → 计算版本号
        → 注入版本号到脚本 → PyInstaller 打包 → 校验产物 → 挂载到 Release
```

> 打包前会把计算出的版本号写入脚本的 `APP_VERSION`，所以每个 exe 都确切知道自己是哪个版本 —— 这是自动更新判断「是否需要升级」的依据。

## 本地开发

```bash
pip install -r requirements.txt

python AIReplay-check.py                  # 打开图形界面
python AIReplay-check.py 20260924         # 静默模式（可用于定时任务）
python AIReplay-check.py --selfcheck      # 界面自检
python AIReplay-check.py --check-update   # 只检测更新并打印结果
python AIReplay-check.py --version        # 打印当前版本号
```

## 本地打包 exe

```bash
pip install pyinstaller
pyinstaller --noconsole --onefile --clean --collect-all customtkinter --name "AIReplay" AIReplay-check.py
```

产物位于 `dist/AIReplay.exe`。

## 目录结构

```
AIReplay-check.py                     # 主程序（图形界面 + 静默模式）
license_guard.py                      # 授权校验 SDK（随程序一起分发）
requirements.txt                      # 运行 / 打包依赖
.github/workflows/build-release.yml   # 自动构建并发布 Release
```
