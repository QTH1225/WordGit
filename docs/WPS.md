# WordGit for WPS（Linux / macOS / Windows）

WordGit 的跨平台版本：Rust 本地服务 + WPS 文字加载项。功能与 Windows Word 版相同：提交版本、分支、返稿合并、对比与恢复、通过 GitHub / Gitee 远程协作。稿件历史与 Word 版使用同一种仓库格式，两边可以互相读取。

## 组成

| 部分 | 位置 | 作用 |
| --- | --- | --- |
| `wordgit` | `crates/wordgit` | 本地服务（只监听 127.0.0.1）、安装命令和命令行 |
| `wordgit-core` | `crates/wordgit-core` | 版本历史内核：分块快照、分支、远程同步、恢复记录 |
| WPS 加载项 | `addin-wps/` | 功能区、"批注、修订与提交"侧栏和对话框 |

加载项不能直接运行 Git，所有历史操作都交给本地服务；对比、修订、打开和关闭文档由加载项调用 WPS 自己完成。

## 安装（Linux）

需要：WPS 文字（个人版 12.1 或更新）、git。

```sh
cargo build --release -p wordgit
./target/release/wordgit install --addin addin-wps
```

`wordgit install` 会：

1. 把加载项复制到 `~/.local/share/Kingsoft/wps/jsaddons/WordGit_<版本>/`，并在 `publish.xml` 中登记，保留其他加载项；
2. 添加登录自启动 `~/.config/autostart/wordgit.desktop`；
3. 启动服务。

然后重新打开 WPS 文字。首次载入、以及升级后的第一次载入，WPS 会询问"是否信任 WordGit 加载项"，选择「确定」。不需要调整宏安全级别。

卸载：`wordgit uninstall`。卸载只移除加载项和自启动，稿件历史保留。

WPS 内置浏览器依赖 NSS、ATK、AT-SPI、ALSA 等系统库。桌面发行版一般自带；如果 WordGit 选项卡不出现、WPS 的登录页也显示"页面加载失败"，请安装 `libnss3 libatk1.0-0 libatk-bridge2.0-0 libatspi2.0-0 libasound2`。

## 数据位置

| 系统 | 目录 |
| --- | --- |
| Linux | `~/.local/share/WordGit` |
| macOS | `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/WordGit`（位于 WPS 沙盒内，WPS 才能打开预览和合并稿） |
| Windows | `%LOCALAPPDATA%\WordGit`（与 Word 版共用） |

目录下有 `repositories`（历史仓库）、`sessions`（预览、对比和合并稿）、`logs/wordgit.log` 和 `service.json`（服务端口与口令，仅当前用户可读）。可用环境变量 `WORDGIT_HOME` 改到别处，`WORDGIT_GIT` 指定 git 路径。

## 命令行

```sh
wordgit status 论文.docx            # 当前分支、提交与远程状态（JSON）
wordgit log 论文.docx               # 提交历史
wordgit commit 论文.docx -m "说明"  # 提交
wordgit call branch/create '{"document":"/path/论文.docx","name":"review"}'
```

`wordgit call` 可以执行服务的任何一个接口，便于脚本和排障。

## 与 Word 版的差异

- **合并稿**：WPS 不支持 Word 的三方修订合并，WordGit 用"对比"生成合并稿。只有对方有修改时结果与 Word 版一致；双方都有修改时，合并稿中也会出现"撤销自己修改"的修订，加载项会同时打开一份"当前分支的修改"对照稿，请把与之相反的修订拒绝。
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
