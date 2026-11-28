# 开发与验证

需要 Windows、.NET Framework 4.8、[Rust 工具链](https://rustup.rs/)（编译内核 `wordgit.exe`），以及 Git for Windows 2.51+。真机测试还需要桌面版 Microsoft Word。

本地开发与 CI 使用同一套工具版本，由仓库根目录的版本文件固定：

| 工具 | 版本文件 | 说明 |
|---|---|---|
| Rust | `rust-toolchain.toml` | Windows/macOS 本机在仓库目录运行 `rustup install` 安装；Linux 容器构建时读取；`docker/wps-linux/Dockerfile` 的 `rust:` 镜像需同步 |
| Node.js | `.node-version` | Linux 容器内的 WPS 加载项脚本检查与 `addin-wps/test/updates.js` |
| Python | `.python-version` | Linux 容器、macOS 安装包测试与发布说明提取；容器基于 Debian trixie，构建时会核对其 Python 主次版本 |

升级任一工具时修改对应文件，Rust 升级还需确认 `cargo clippy --workspace --all-targets -- -D warnings` 通过。

## 分工：哪些在 Docker 里，哪些在 Windows 本机

| 在哪里 | 命令 | 产出 / 用途 |
|---|---|---|
| Windows 本机 | `scripts\build.ps1`、`scripts\package.ps1`（见下文） | `bin\WordGit.dll`、`bin\wordgit.exe`、`dist\WordGit-<版本>-Setup.exe` 与 `-windows.zip`，装到本机 Word / WPS 实测 |
| Linux 开发容器 | `scripts\dev.ps1 check`（Windows）或 `scripts/dev.sh check`（Linux/macOS） | 编码检查、WPS 加载项脚本检查、clippy、内核测试、Linux 安装包安装/升级/卸载测试，与 CI 的 `linux` 任务完全相同 |
| Linux 开发容器 | `scripts\dev.ps1 package` | `dist/WordGit-<版本>-linux-x86_64.tar.gz`（musl 静态版，与发布流程相同） |
| Linux WPS 容器 | 见 [docker/wps-linux/README.md](../docker/wps-linux/README.md) | 在 Linux WPS 文字中端到端实测 |
| macOS 本机 | 见 [MACOS.md](MACOS.md) | macOS 安装包 |

Word 加载项依赖 Windows 自带的 .NET Framework 编译器和桌面版 Word，内核 `wordgit.exe` 需要 MSVC 工具链，这些只能在 Windows 本机构建和测试。

### Linux 开发容器

需要 Docker（Windows 上为 Docker Desktop 的 Linux 容器模式）。镜像由 `docker/dev/Dockerfile` 在本地构建，不推送到任何镜像仓库；标签取自 Dockerfile 与三个版本文件的哈希，版本文件一改，下次运行自动重建并清理旧镜像。

```powershell
.\scripts\dev.ps1 check                         # 与 CI 相同的 Linux 检查
.\scripts\dev.ps1 cargo test -p wordgit-core    # 在容器中运行任意命令
.\scripts\dev.ps1                               # 进入容器 shell
```

源码以 `/work` 挂载进容器；构建产物、crate 缓存放在名为 `wordgit-dev-cache` 的 Docker 卷中，不会与 Windows 本机的 `target\` 混用。需要清空缓存时运行 `docker volume rm wordgit-dev-cache`。

国内网络构建镜像较慢时，可在运行前设置下载镜像（只影响下载来源，不影响版本）：

```powershell
$env:WORDGIT_APT_MIRROR = 'https://mirrors.huaweicloud.com'
$env:WORDGIT_NODE_MIRROR = 'https://registry.npmmirror.com/-/binary/node'
$env:WORDGIT_DOCKER_REGISTRY = '<Docker Hub 镜像地址>/library/'   # 可选，基础镜像前缀
```

普通用户应使用预编译发布包。以下步骤供开发者使用，PowerShell 5.1 或 PowerShell 7 均可。

## 构建

在项目根目录打开 PowerShell：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Test -UiTest
```

输出界面插件 `bin/WordGit.dll` 和内核 `bin/wordgit.exe`。脚本先用系统 .NET Framework 编译器（C# 5）编译界面，不下载 NuGet 依赖，再用 `cargo build --release -p wordgit` 编译内核并复制过来；只想重编界面时加 `-NoRust`。

每个 C# 测试程序都由 src 全部源码、`tests/TestSupport.cs` 共享辅助与对应的测试文件一起编译。辅助方法会自行拉起一个临时内核（`WORDGIT_SERVICE_EXE` 指向 `bin/wordgit.exe`，数据目录用临时的 `WORDGIT_HOME`），测试因此走与插件相同的本地服务接口，不另设仅供测试的内存路径。请求的测试程序会全部运行，结束后汇总列出失败项。

| 参数 | 测试 | 需要 |
| --- | --- | --- |
| `-Test` | UpdatesSmoke（更新检查） | — |
| `-UiTest` | UiSmoke（功能区、对话框、侧栏布局与交互） | 桌面会话 |
| `-PaneTest` | PaneProbe（侧栏首帧先于读取、空历史状态行不折行） | 桌面会话 |
| `-KernelTest` | KernelSmoke（经内核完成提交、预览、恢复、分支合并与多版本合并） | 桌面版 Word |
| `cargo test --workspace --release` | 内核单元与集成测试（存储格式、分支、返稿冲突、崩溃恢复、双客户端远程、服务安全） | Rust 工具链 |

GitHub Actions 在每次推送和 PR 时运行 `-Test`、`-UiTest` 并生成发布包，Linux 部分通过 `scripts/dev.sh check` 在同一个开发容器中检查，Windows 与 macOS 另在本机运行内核测试（`.github/workflows/ci.yml`）；CodeQL 每周及推送时扫描 C# 源码与工作流。

真机测试会自己启动隐藏的 Word 实例，新建独立测试文档与会话：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 -KernelTest
```

## 从源码安装

构建完成、保存并关闭 Word 后：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\install.ps1
```

只检查环境和文件：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\install.ps1 -CheckOnly
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\install.ps1 -Plan
```

-CheckOnly 与 -Plan 均不安装；后者输出 JSON 形式的安装计划。

## 生成发布包

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\package.ps1 -Test
```

也可增加 -KernelTest，执行 Word 真机的提交、预览、恢复与合并测试。打包前会核对 README 与 docs/RELEASING.md 中的版本号是否与插件一致。发布文件位于 dist/：

- WordGit-版本-Setup.exe：内嵌预编译插件、内核与安装文件的独立安装维护程序。
- WordGit-版本-windows.zip：含相同安装程序、脚本、文档、预编译 DLL 与 `addin-wps/`。
- SHA256SUMS.txt：EXE 和 ZIP 的 SHA-256。

-Test 验证核心历史、混合存储格式、双客户端本地远程同步、界面、内嵌资源、Windows PowerShell 5.1、安装与升级的数据保留及卸载。Git for Windows 的远程辅助进程需要正常 Windows 进程权限，受限沙箱可能阻止运行。安装测试使用隔离目录和临时注册表子项，不卸载当前用户正在使用的插件。

## 存储与注册

运行时源码分两层：`crates/wordgit-core` 是版本历史内核（Git 对象、分块快照、分支、返稿合并、远程同步、整理与崩溃恢复），`crates/wordgit` 把它包成本地 HTTP 服务并负责安装；`src/` 只剩 Word 界面、Word COM 与 `ServiceClient`/`ServiceModels` 两个适配层，`addin-wps/` 是 WPS 的 JS 加载项，安装维护程序位于 `packaging/Setup.cs`。Word 侧采用 .NET Framework COM 接口，无 VSTO 依赖。

服务只监听 `127.0.0.1`，端口与 48 位十六进制口令写在 `%LOCALAPPDATA%\WordGit\service.json`（退出即删除）。除 `/api/v1/ping` 外所有请求都要带 `x-wordgit-token`，并校验 `Host` 与 `Origin`，因此浏览器页面无法借用本机服务。加载项启动时读取该文件；口令对不上会提示重启编辑器。

WordGit.Connect 和 WordGit.HistoryPane 注册到当前用户的 COM 类和 Word 加载项项下，安装器根据 WINWORD.EXE 的实际架构选择 32/64 位注册表视图。CLSID 与 ProgID 同时写在源码属性和 scripts/install.ps1 开头的变量中，安装测试会核对两者一致。卸载不处理历史仓库或审阅稿。内核用 `HKCU\...\Run` 的 `WordGit` 值登录自启，WPS 加载项由 `wordgit install --addin` 写入 `jsaddons` 与 `publish.xml`。

所有 git.exe 调用都在 `crates/wordgit-core/src/git.rs`：清除继承的 GIT_* 环境变量、屏蔽全局与系统配置（网络操作除外，以使用用户的凭据和代理）、`core.longpaths`、`core.fsync=committed`，并支持超时取消与整棵进程树结束。新增 Git 调用时不要绕过它直接启动进程。

替换用户的当前稿只能走两段式：加载项先调用 `branch/switch`、`version/restore`、`merge/finish` 或 `remote/pull` 拿到计划（计划持有该文档的锁，10 分钟过期），关掉文档后调用 `plan/apply`；内核 `review::replace_working_copy` 先写 journal 意图、再用备份替换，失败时回滚，Word 中途退出后启动时按 journal 恢复。放弃计划必须调用 `plan/cancel`，否则文档会被锁住。Word 线程上的异步入口在第一次 await 前调用 `UiThread.EnsureContext()`。

仓库支持整包 document.docx、wordgit.meta 拆包格式、wordgit.snapshot 原始字节分块格式，读取实现只在 `crates/wordgit-core`。三种格式必须持续保留读取测试（`reads_retired_storage_formats`）；更换存储格式前，应验证真实旧仓库、混合格式、压缩前后取回，以及分支与提交关系。

不要将用户 DOCX、repositories、sessions、恢复文件或本机环境信息放入发布包。
