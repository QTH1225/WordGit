# WordGit for WPS（Linux / macOS / Windows）

WPS 文字通过 JS 加载项连接 WordGit 的 Rust 本地服务，支持提交版本、分支、返稿合并、对比、恢复和 GitHub / Gitee 远程协作。Word 和 WPS 使用同一种历史仓库格式。

## 组成

| 部分 | 位置 | 作用 |
| --- | --- | --- |
| `wordgit` | `crates/wordgit` | 本地服务（只监听 127.0.0.1）、安装命令和命令行 |
| `wordgit-core` | `crates/wordgit-core` | 版本历史内核：分块快照、分支、远程同步、恢复记录 |
| WPS 加载项 | `addin-wps/` | 功能区、"批注、修订与提交"侧栏和对话框 |

加载项不能直接运行 Git，所有历史操作都交给本地服务；对比、修订、打开和关闭文档由加载项调用 WPS 自己完成。

## 安装（Windows）

若同时安装了 Microsoft Word，运行 WordGit 的 Windows 图形安装程序即可部署两个加载项。只使用 WPS 文字时，将 ZIP 安装包解压到长期保留的目录，安装 Git for Windows 2.51 或更新版，并在解压目录运行：

```powershell
.\bin\wordgit.exe install --addin .\addin-wps
```

命令会部署 WPS 加载项、注册登录自启并启动内核。重新打开 WPS 文字完成加载。卸载时在同一目录运行 `.\bin\wordgit.exe uninstall`，已有历史保留。

## 安装（Linux）

需要 WPS 文字个人版 12.1 或更新版本，以及 Git。Linux x86_64 可从 GitHub Release 下载压缩包，解压后运行 `./install.sh`。源码构建使用以下命令。

```sh
cargo build --release -p wordgit
./target/release/wordgit install --addin addin-wps
```

`wordgit install` 会：

1. 把加载项复制到 `~/.local/share/Kingsoft/wps/jsaddons/WordGit_<版本>/`，并在 `publish.xml` 中登记，保留其他加载项；
2. 将内核安装到 `~/.local/share/WordGit/app/wordgit`，并将登录自启指向该路径。安装完成后，解压目录可以删除；
3. 关闭旧的 WordGit 服务并启动新版内核。升级前请先退出 WPS，再运行新版压缩包的 `./install.sh`。

然后重新打开 WPS 文字。首次载入、以及升级后的第一次载入，WPS 会询问"是否信任 WordGit 加载项"，选择「确定」。不需要调整宏安全级别。

卸载可运行压缩包中的 `./uninstall.sh`，或运行 `~/.local/share/WordGit/app/wordgit uninstall`。卸载会停止服务，移除内核、加载项和自启记录，保留已有历史与审阅稿。

WPS 内置浏览器依赖 NSS、ATK、AT-SPI、ALSA 等系统库。桌面发行版一般自带；如果 WordGit 选项卡不出现、WPS 的登录页也显示"页面加载失败"，请安装 `libnss3 libatk1.0-0 libatk-bridge2.0-0 libatspi2.0-0 libasound2`。

## 安装（macOS）

需要 WPS Office for Mac（≥ 6.15，建议 12.1 或更新版本）以及 Git（`xcode-select --install` 或 Homebrew Git）。详见 [macOS 安装与使用说明](MACOS.md)。

下载并解压 `WordGit-<版本>-macos-<架构>.tar.gz` 后，双击 `Install.command` 或在终端运行 `./install.sh`。从源码构建：

```sh
cargo build --release -p wordgit
./target/release/wordgit install --addin addin-wps
```

`wordgit install` 会在 macOS 上：

1. 将加载项复制到 `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/WordGit_<版本>/` 并更新 `publish.xml`；
2. 将内核安装到 `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/WordGit/app/wordgit`；
3. 写入并加载当前用户的 LaunchAgent `~/Library/LaunchAgents/com.qth1225.wordgit.plist`，关闭旧服务并启动新版内核。

卸载可运行 `./uninstall.sh`（或双击 `Uninstall.command`），或执行 `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/WordGit/app/wordgit uninstall`，已有历史与审阅稿保留。

## 数据位置

| 系统 | 目录 |
| --- | --- |
| Linux | `~/.local/share/WordGit` |
| macOS | `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/WordGit`（位于 WPS 沙盒内，WPS 才能打开预览和合并稿） |
| Windows | `%LOCALAPPDATA%\WordGit`（与 Word 版共用） |

通过 **WordGit → 数据目录** 查看或迁移 `repositories` 历史仓库。Windows 安装时也可选择位置。Linux 首次安装可执行 `./install.sh --data-directory /完整/路径`。

`sessions`（预览、对比和合并稿）、日志以及 `service.json` 仍由服务保存在上表的运行目录中，以便编辑器继续访问已打开的审阅稿。`WORDGIT_HOME` 可指定运行目录，`WORDGIT_GIT` 可指定 Git 路径。

## 命令行

```sh
wordgit status 论文.docx            # 当前分支、提交与远程状态（JSON）
wordgit log 论文.docx               # 提交历史
wordgit commit 论文.docx -m "说明"  # 提交
wordgit call branch/create '{"document":"/path/论文.docx","name":"review"}'
```

`wordgit call` 可以执行服务的任何一个接口，便于脚本和排障。

## 与 Word 版的差异

- **合并稿**：WPS 使用自身的对比功能生成修订。返稿合并以选定的送审版本为基准。远程分歧会提示选择保留本地或远端的冲突内容。部分本地分支合并会另外打开「当前分支的修改」对照稿，供审阅时参考。
- **切换分支、恢复、拉取、完成合并**：加载项先关闭当前稿，由服务替换文件并更新分支，再重新打开。替换前会写入恢复记录，中途崩溃后服务下次启动时自动还原或完成。
- **克隆稿件**：Windows 上会给克隆的稿件加"来自互联网"标记；Linux 和 macOS 没有这个机制。

## 开发与测试

```sh
cargo test --workspace --release             # 单元测试、真实 git 集成测试、服务 HTTP 测试
```

Linux WPS 端到端测试在 Docker 中运行，见 [docker/wps-linux/README.md](../docker/wps-linux/README.md)：

```sh
PROBE_MODE=e2e docker compose -f docker/wps-linux/compose.yaml up -d
```

它在真实的 WPS 个人版中载入加载项，依次执行提交、预览、对比、分支切换、分支合并、恢复、另存、返稿导入与合并、关联、推送、克隆、快进拉取、分叉拉取和侧栏渲染，结果写入 `docker/wps-linux/out/e2e-report.jsonl`。

编写加载项时注意：WPS 默认的整合窗口模式下，在一次同步执行中连续打开、关闭多个文档会让 WPS 崩溃；文档操作之间一律 `await` 让出事件循环（`addin-wps/js/core.js` 的 `WG.open` / `WG.close` 已经这样做）。
