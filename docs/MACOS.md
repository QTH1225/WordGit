# WordGit for macOS（WPS 文字）

本指南说明如何在 macOS（Apple Silicon M 系列芯片与 Intel 芯片）上安装、构建、升级和卸载 WordGit。macOS 版本与 Windows / Linux 版本共用同一套 Rust 内核、历史仓库格式和 WPS JS 加载项功能区，完整支持提交、提交图、分支、返稿合并、修订对比、历史恢复以及 GitHub / Gitee 远程协作。

## 环境要求

| 项目 | 要求 |
| --- | --- |
| 系统 | macOS 12+（Apple Silicon arm64 / Intel x86_64） |
| 编辑器 | WPS Office for Mac（≥ 6.15，建议 12.1 或更新版本） |
| Git | Xcode 命令行工具（`xcode-select --install`）或 Homebrew Git（`brew install git`） |

> [!NOTE]
> macOS 上的 WPS Office 运行在 App Sandbox 沙盒内，因此 WordGit 将本地运行目录（`service.json`、会话预览/合并稿 `sessions`、日志 `logs` 以及内核二进制 `app/wordgit`）放在 WPS 沙盒容器内：
> `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/WordGit`
> 历史仓库（`repositories`）默认也位于此处，可在 WPS 的 **WordGit → 数据目录** 中迁移到任意本地目录。

## 安装方式

### 方式一：使用打包好的 macOS 发布包（推荐）

1. 解压 `WordGit-<版本>-macos-<架构>.tar.gz`（例如 `WordGit-0.14.0-macos-arm64.tar.gz`）。
2. 保存文档并完全退出 WPS Office（`⌘ + Q`）。
3. 双击解压目录中的 **`Install.command`**，或在终端运行：

   ```sh
   ./install.sh
   ```

   首次安装时如需指定历史仓库目录，可附加参数：

   ```sh
   ./install.sh --data-directory "/完整/路径"
   ```

4. 重新打开 WPS Office。首次载入（或升级后首次载入）时，WPS 会弹出“是否信任 WordGit 加载项”，请选择 **「确定」**。顶部出现 **WordGit** 选项卡即安装成功。

安装程序会自动完成以下步骤：

1. 将 JS 加载项复制到 `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/WordGit_<版本>/`，并在 `publish.xml` 中注册（保留其他已有加载项）；
2. 将本地内核 `wordgit` 安装到 `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/WordGit/app/wordgit`（安装完成后可删除解压目录）；
3. 写入并加载当前用户的 LaunchAgent：`~/Library/LaunchAgents/com.qth1225.wordgit.plist`，设置登录自启与进程保活；
4. 优雅停止旧版服务、启动新版服务并校验本地 HTTP 接口（`127.0.0.1:47821`）就绪。

### 方式二：从源码构建与打包

需要安装 Rust 工具链（`brew install rust` 或 [rustup.rs](https://rustup.rs/)）：

```sh
# 1. 编译 macOS 原生内核
cargo build --locked --release -p wordgit

# 2. 直接安装到本机 WPS
./target/release/wordgit install --addin addin-wps

# 或生成可分发的 macOS 发布包（输出到 dist/）
bash scripts/package-macos.sh target/release/wordgit dist
```

如需构建同时支持 Apple Silicon 与 Intel 的通用二进制（Universal Binary）：

```sh
rustup target add aarch64-apple-darwin x86_64-apple-darwin
cargo build --locked --release --target aarch64-apple-darwin -p wordgit
cargo build --locked --release --target x86_64-apple-darwin -p wordgit
lipo -create \
  target/aarch64-apple-darwin/release/wordgit \
  target/x86_64-apple-darwin/release/wordgit \
  -output target/release/wordgit-universal
bash scripts/package-macos.sh target/release/wordgit-universal dist universal
```

## 升级与卸载

- **升级**：先退出 WPS Office，然后运行新版包中的 `./install.sh`（或双击 `Install.command`）。安装程序会自动停止旧服务、替换内核与加载项并重启服务，已有的提交、分支、审阅稿与自定义数据目录全部保留。
- **卸载**：双击 `Uninstall.command`，或运行 `./uninstall.sh` /：

  ```sh
  ~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/WordGit/app/wordgit uninstall
  ```

  卸载会停止并移除 LaunchAgent、删除内核与 WPS 加载项注册，但保留已有稿件历史（`repositories`）和审阅稿。

## 常见问题排查

- **macOS 提示“无法打开，因为无法验证开发者”**：
  使用 `Install.command` 会自动清除解压目录的隔离属性；如手动运行，可执行 `xattr -dr com.apple.quarantine <解压目录>` 后再运行 `./install.sh`。
- **服务未连接 / 端口占用**：
  在终端检查服务状态：
  ```sh
  curl http://127.0.0.1:47821/api/v1/ping
  launchctl print "gui/$(id -u)/com.qth1225.wordgit"
  ```
- **快捷键差异**：
  在 macOS 上支持使用 `⌘ + Enter` 提交版本说明，以及按住 `⌘` 在提交图中同时选择两个历史版本进行对比或合并。
- **查看错误日志**：
  `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/WordGit/logs/wordgit.log`
