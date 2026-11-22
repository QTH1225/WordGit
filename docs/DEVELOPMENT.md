# 开发与验证

需要 Windows、.NET Framework 4.8，以及用于功能测试的 Git for Windows 2.51+。Word 原生测试还需要桌面版 Microsoft Word。

普通用户应使用预编译发布包。以下步骤供开发者使用，PowerShell 5.1 或 PowerShell 7 均可。

## 构建

在项目根目录打开 PowerShell：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Test -UiTest
```

输出插件 DLL 到 bin/WordGit.dll。脚本调用系统 .NET Framework 编译器，不下载 NuGet 依赖。

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

也可增加 -WordTest，执行 Word 原生比较、修订、合并和恢复测试。发布文件位于 dist/：

- WordGit-版本-Setup.exe：内嵌预编译插件与安装文件的独立安装维护程序。
- WordGit-版本-windows.zip：含相同安装程序、脚本、文档与预编译 DLL。
- SHA256SUMS.txt：EXE 和 ZIP 的 SHA-256。

-Test 验证核心历史、混合存储格式、双客户端本地远程同步、界面、内嵌资源、Windows PowerShell 5.1、安装与升级的数据保留及卸载。Git for Windows 的远程辅助进程需要正常 Windows 进程权限，受限沙箱可能阻止运行。安装测试使用隔离目录和临时注册表子项，不卸载当前用户正在使用的插件。

## 存储与注册

运行时源码位于 src/；安装维护程序位于 packaging/Setup.cs。插件采用 .NET Framework COM 接口，无 VSTO 依赖。

WordGit.Connect 和 WordGit.HistoryPane 注册到当前用户的 COM 类和 Word 加载项项下，安装器根据 WINWORD.EXE 的实际架构选择 32/64 位注册表视图。卸载不处理历史仓库或审阅稿。

仓库支持整包 document.docx、wordgit.meta 拆包格式、wordgit.snapshot 原始字节分块格式。三种格式必须持续保留读取测试；更换存储格式前，应验证真实旧仓库、混合格式、压缩前后取回，以及分支与提交关系。

不要将用户 DOCX、repositories、sessions、恢复文件或本机环境信息放入发布包。
