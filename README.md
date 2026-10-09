# AIReplay-check

Replay 核查一体化工具：一次运行完成「后台抓取 → CSV 导出 → 筛选 / 排序 / 去重 → xlsx 输出」，带图形操作界面。

[![Build EXE & Publish Release](https://github.com/Serein1202/AIReplay-check/actions/workflows/build-release.yml/badge.svg)](https://github.com/Serein1202/AIReplay-check/actions/workflows/build-release.yml)

## 下载使用

到 **[Releases](https://github.com/Serein1202/AIReplay-check/releases/latest)** 页面下载 `AIReplay-<版本号>.exe`（例如 `AIReplay-v1.0.1.exe`），双击运行即可，**无需安装 Python**。

首次运行需要在弹窗中登录以获取 Token。Token 会缓存在 exe 同目录的 `.amis_token.json`，**内含凭据，请勿外传或提交到仓库**（已在 `.gitignore` 中忽略）。

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
requirements.txt                      # 运行 / 打包依赖
.github/workflows/build-release.yml   # 自动构建并发布 Release
```
