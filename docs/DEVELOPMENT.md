# 开发与验证

需要 Windows、.NET Framework 4.8、[Rust 工具链](https://rustup.rs/)（编译内核 `wordgit.exe`），以及 Git for Windows 2.51+。真机测试还需要桌面版 Microsoft Word。

本地开发与 CI 使用同一套工具版本，由仓库根目录的版本文件固定：

| 工具 | 版本文件 | 说明 |
|---|---|---|
| Rust | `rust-toolchain.toml` | 在仓库目录运行 `rustup install` 安装；`docker/wps-linux/Dockerfile` 的 `rust:` 镜像需同步 |
| Node.js | `.node-version` | WPS 加载项脚本检查与 `addin-wps/test/updates.js` |
| Python | `.python-version` | 发布说明提取与 Linux/macOS 安装包测试 |

升级任一工具时修改对应文件，Rust 升级还需确认 `cargo clippy --workspace --all-targets -- -D warnings` 通过。

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

GitHub Actions 在每次推送和 PR 时运行 `-Test`、`-UiTest` 并生成发布包（`.github/workflows/ci.yml`）；CodeQL 每周及推送时扫描 C# 源码与工作流。

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
