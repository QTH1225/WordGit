# 开发与验证

需要 Windows、.NET Framework 4.8，以及用于功能测试的 Git for Windows 2.51+。Word 原生测试还需要桌面版 Microsoft Word。

普通用户应使用预编译发布包。以下步骤供开发者使用，PowerShell 5.1 或 PowerShell 7 均可。

## 构建

在项目根目录打开 PowerShell：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Test -UiTest
```

输出插件 DLL 到 bin/WordGit.dll。脚本调用系统 .NET Framework 编译器（C# 5），不下载 NuGet 依赖。

每个测试程序都由 src 全部源码、`tests/TestSupport.cs` 共享辅助与对应的测试文件一起编译。辅助方法让测试可以直接使用 byte[]，但内部统一走插件实际使用的文件与流式接口（`CommitFile`、`ExtractTo`、`ReviewPaths`、`ReplaceWorkingCopy`），不另设仅供测试的内存路径。请求的测试程序会全部运行，结束后汇总列出失败项。

| 参数 | 测试 | 需要 |
| --- | --- | --- |
| `-Test` | Smoke（历史、格式兼容、维护、恢复记录）、UpdatesSmoke、RemoteSmoke（双客户端本地远程） | Git for Windows |
| `-UiTest` | UiSmoke（功能区、对话框、侧栏布局） | 桌面会话 |
| `-WordTest` | WordSmoke（原生比较、修订、合并、恢复、远程协作） | 桌面版 Word |

GitHub Actions 在每次推送和 PR 时运行 `-Test`、`-UiTest` 并生成发布包（`.github/workflows/ci.yml`）；CodeQL 每周及推送时扫描 C# 源码与工作流。

Word 原生功能测试会新建独立会话和测试文档（包含双客户端远程协作场景）：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -WordTest
```

## 从源码安装

构建完成、保存并关闭 Word 后：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
```

只检查环境和文件：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -CheckOnly
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -Plan
```

-CheckOnly 与 -Plan 均不安装；后者输出 JSON 形式的安装计划。

## 生成发布包

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\package.ps1 -Test
```

也可增加 -WordTest，执行 Word 原生比较、修订、合并和恢复测试。打包前会核对 README 与 docs/RELEASING.md 中的版本号是否与插件一致。发布文件位于 dist/：

- WordGit-版本-Setup.exe：内嵌预编译插件与安装文件的独立安装维护程序。
- WordGit-版本-windows.zip：含相同安装程序、脚本、文档与预编译 DLL。
- SHA256SUMS.txt：EXE 和 ZIP 的 SHA-256。

-Test 验证核心历史、混合存储格式、双客户端本地远程同步、界面、内嵌资源、Windows PowerShell 5.1、安装与升级的数据保留及卸载。Git for Windows 的远程辅助进程需要正常 Windows 进程权限，受限沙箱可能阻止运行。安装测试使用隔离目录和临时注册表子项，不卸载当前用户正在使用的插件。

## 存储与注册

运行时源码位于 src/；安装维护程序位于 packaging/Setup.cs。插件采用 .NET Framework COM 接口，无 VSTO 依赖。

WordGit.Connect 和 WordGit.HistoryPane 注册到当前用户的 COM 类和 Word 加载项项下，安装器根据 WINWORD.EXE 的实际架构选择 32/64 位注册表视图。CLSID 与 ProgID 同时写在源码属性和 install.ps1 开头的变量中，安装测试会核对两者一致。卸载不处理历史仓库或审阅稿。

所有 git.exe 调用经过 `GitProcess`：清除继承的 GIT_* 环境变量、屏蔽全局与系统配置（网络操作除外，以使用用户的凭据和代理）、开启 `core.fsync=committed`，并把进程放进作业对象，超时或取消时结束整棵进程树。新增 Git 调用时不要绕过它直接启动进程。

替换用户的当前稿只能通过 `WordDocuments.ReplaceWorkingCopy`：它写入 `WorkingCopyJournal` 意图记录、用 `File.Replace` 保留备份，并在发布引用失败时回滚。Word 线程上的异步入口在第一次 await 前调用 `UiThread.EnsureContext()`。

仓库支持整包 document.docx、wordgit.meta 拆包格式、wordgit.snapshot 原始字节分块格式。三种格式必须持续保留读取测试；更换存储格式前，应验证真实旧仓库、混合格式、压缩前后取回，以及分支与提交关系。

不要将用户 DOCX、repositories、sessions、恢复文件或本机环境信息放入发布包。
