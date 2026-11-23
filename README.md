# WordGit 0.11.2

在 Windows 桌面版 Word 和 WPS 文字中提交版本、管理分支、查看历史，并通过 GitHub 同步稿件。文档预览、版本对比和修订处理使用当前编辑器的原生功能。

作者：[QTH1225](https://github.com/QTH1225) · [项目主页](https://github.com/QTH1225/WordGit) · [发布记录](https://github.com/QTH1225/WordGit/releases) · [问题反馈](https://github.com/QTH1225/WordGit/issues) · [GPL-3.0-only 许可证](LICENSE)

## 安装与升级

| 环境 | 要求 |
| --- | --- |
| 系统 | Windows 10 / 11 |
| 编辑器 | Windows 桌面版 Microsoft Word 或 WPS 文字 |
| Git | [Git for Windows 2.51+](https://git-scm.com/install/windows) |
| 运行时 | [.NET Framework 4.8+](https://dotnet.microsoft.com/download/dotnet-framework/net48) |

1. 准备 `WordGit-0.11.2-Setup.exe` 安装包。
2. 保存文档，退出 Word 和 WPS 文字。
3. 运行安装程序，点击 **安装 / 更新**。
4. 打开 Word 或 WPS 文字，进入顶部 **WordGit** 选项卡。

ZIP 包的用法相同：完整解压 `WordGit-0.11.2-windows.zip`，运行其中的 `WordGit.Setup.exe`。也可以双击 `Install.cmd` 安装、`Check.cmd` 检查环境。

安装作用于当前 Windows 用户。升级时直接运行新版安装程序，已有提交、分支和审阅稿继续保留。卸载可从 Windows“已安装的应用”或安装程序进入，历史数据也会保留。

### WPS 文字兼容范围

WPS 文字可独立提交、预览、恢复、对比和处理现有修订。**返稿合并**入口在 WPS 中可将原始返稿保存为分支，保留批注和修订。自动生成返稿、历史版本或分支的修订合并稿，以及分叉后的拉取合并，目前需要在 Word 中操作。

如未出现 **WordGit** 选项卡，在 WPS 文字的 **开发工具 → COM 加载项** 查看 WordGit 是否启用。在 Word 中的加载步骤见下文。

## 在 Word 中载入插件

WordGit 是 Word 的 COM 加载项，**不需要在 Word 里手动添加文件**：安装程序完成注册后，Word 每次启动都会自动载入它。

### 载入方式

安装程序只写当前用户（HKCU）的注册表，不需要管理员权限：

| 注册项 | 作用 |
| --- | --- |
| `Software\Classes\CLSID\{C5707F76-F46F-4A11-800B-18E7D98377D3}` 与 `WordGit.Connect` | 把插件登记为 COM 组件，指向 `%LOCALAPPDATA%\WordGit\app\WordGit.<版本>.<哈希>.dll`。 |
| `Software\Classes\CLSID\{F55EC3C1-B725-4ED8-9CA6-ED51D07C3361}` 与 `WordGit.HistoryPane` | 右侧“批注、修订与提交”窗格的控件。 |
| `Software\Microsoft\Office\Word\Addins\WordGit.Connect` | 告诉 Word 有这个加载项；`LoadBehavior = 3` 表示 Word 启动时自动载入。 |

安装程序会检测 Word 和 WPS 文字的实际位数，再写入相应的注册表视图。WPS 的加载项登记在 `Software\Kingsoft\Office\WPS\Addins\WordGit.Connect`，两个编辑器使用同一个安装包。

### 确认已经载入

1. 安装完成后**完全关闭 Word 再重新打开**。仍在运行的 Word 不会载入新插件，可在任务管理器确认没有残留的 `WINWORD.EXE`。
2. 顶部出现 **WordGit** 选项卡即表示载入成功。
3. 也可以在 Word 中打开 **文件 → 选项 → 加载项**，“活动应用程序加载项”列表里应有 **WordGit**，类型为“COM 加载项”。

### 没有看到 WordGit 选项卡时

| 情况 | 处理 |
| --- | --- |
| 加载项被取消勾选 | **文件 → 选项 → 加载项**，底部“管理”选择 **COM 加载项** → **转到**，勾选 **WordGit** → **确定**。 |
| 被 Word 自动禁用（Word 启动时崩溃或加载过慢后会这样做） | 同一页面“管理”选择 **禁用项目** → **转到**，选中 WordGit → **启用**，然后重启 Word，再按上一行勾选。 |
| 环境不完整 | 双击 ZIP 包中的 `Check.cmd`，或在安装程序中点击“检查环境”，确认 Word、Git for Windows 与 .NET Framework 4.8；缺什么装什么，再重新运行安装程序。 |
| 单位电脑开启了“要求应用程序加载项由受信任的发布者签名” | 当前安装包未做代码签名，会被该策略拦截（位置：**文件 → 选项 → 信任中心 → 信任中心设置 → 加载项**）。这项通常由管理员统一设置，需请管理员放行。 |
| 以安全模式启动 Word（如 `winword /safe`） | 安全模式不载入任何加载项，请正常启动 Word。 |
| 刚升级完 | 确认已关闭所有 Word 窗口后再运行新版安装程序；仍不出现时重新运行一次 **安装 / 更新**。 |

不要在“COM 加载项”对话框中用 **添加** 按钮浏览 WordGit 的 DLL：.NET 编写的加载项需要上面的注册表登记，直接添加 DLL 文件无效。

### 暂时停用与卸载

- 暂时停用：在“COM 加载项”对话框中取消勾选 WordGit，Word 下次启动不再载入；重新勾选即可恢复。
- 卸载：从 Windows“已安装的应用”找到 WordGit，或运行安装程序点击 **卸载**。卸载只删除插件文件和上述注册项，稿件、`repositories` 历史、`sessions` 审阅稿和日志都会保留。

### 从源码载入

开发者不需要安装包，在项目根目录依次运行：

```powershell
# 1. 编译插件，生成 bin\WordGit.dll
pwsh -NoProfile -ExecutionPolicy Bypass -File .\build.ps1

# 2. 关闭所有 Word 窗口后，注册到当前用户（与安装程序写入相同的注册项）
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
```

`install.ps1` 面向 Windows PowerShell 5.1，用 `powershell.exe` 运行。重新编译后再运行一次 `install.ps1`，Word 下次启动就会载入新版本；`install.ps1 -Uninstall` 可撤销注册。

## 快速开始

1. 将稿件保存为本地 `.docx`，例如 `D:\Manuscripts\paper.docx`。WordGit 需要能直接读取的本地文件：OneDrive / SharePoint 中以网址打开的稿件，请打开同步到本机文件夹里的那份副本。需要切换分支、原地恢复或合并时，使用 OneDrive 等同步目录之外的本地文件夹。
2. 点击 **WordGit → 提交版本**，填写提交标题，按需补充详细内容。点击“提交”或按 Ctrl + Enter，插件保存当前稿并创建提交。
3. 点击 **提交图** 打开右侧窗格，选中历史版本后点击 **预览** 或 **对比当前稿**。
4. 继续修改，在需要保留进展时再次提交。

## 功能区

### 本地提交

| 按钮 | 用法 |
| --- | --- |
| 提交版本 | 保存当前稿，填写说明，创建一条本地提交。 |
| 提交图 | 打开右侧“批注、修订与提交”窗格，记录在后台加载，可继续编辑正文。 |
| 预览版本 | 选择一个历史提交，在 Word 中只读打开。 |

### 分支

| 按钮 | 用法 |
| --- | --- |
| 新建分支 | 选择起始提交，输入分支名，用于另一个写作方向或审阅版本。 |
| 切换分支 | 选择目标分支，将当前稿切换到该分支的版本；有未提交修改时先提交。 |
| 合并分支 | 选择要合入当前分支的来源分支，在 Word 修订稿中确认修改，再提交合并结果。 |
| 删除分支 | 删除已经合入当前分支的其他本地分支；当前分支不可删除。 |

### 比较与审阅

| 按钮 | 用法 |
| --- | --- |
| 返稿合并 | 在同一个窗口选择返稿文件和送审版本，核对合入分支，生成 Word 修订合并稿。可将原始返稿保存为独立分支。 |
| 对比版本 | 选择一版与当前稿对比；按住 Ctrl 选择两版，可以对比两个历史版本。 |

### 远程协作

一份稿件关联一个 GitHub 仓库，协作者通过 WordGit 交换已提交的版本。

| 按钮 | 用法 |
| --- | --- |
| 关联仓库 | 为当前稿件填写 GitHub 仓库地址和远程分支，连接已有本地历史。 |
| 克隆稿件 | 从 WordGit 仓库取回稿件与历史，保存到一个新的本地 DOCX 路径。 |
| 获取 | 下载远程提交和分支，查看对方进展，当前正文保持原样。 |
| 拉取 | 将远程进展合入当前分支；双方各有新增提交时，进入 Word 修订合并流程。 |
| 推送 | 上传当前分支已经提交的版本。共享其他分支时，切换到该分支再推送。 |

### 取回版本

| 按钮 | 用法 |
| --- | --- |
| 另存旧版 | 将所选历史版本保存为一份新的 DOCX。 |
| 恢复旧版 | 先归档当前稿，再恢复所选版本，并创建一条新提交。 |

### 关于与更新

| 按钮 | 用法 |
| --- | --- |
| WordGit 版本号 | 查看版本、作者、项目主页和反馈链接。 |
| 检查更新 | 查询最新正式发布版本，打开发布页下载安装。 |

## 右侧任务窗格

上半区用于审阅，下半区用于管理版本。拖动中间分隔条可以调整两部分的大小。

| 区域 | 用法 |
| --- | --- |
| 批注与修订列表 | 查看类型、作者、页码和内容。点击一条记录跳转到正文，使用“全部 / 批注 / 修订”筛选。 |
| 接受 / 拒绝 | 选中一条修订后直接处理，随后自动跳到下一条。可用 Word 的 Ctrl+Z 撤销。 |
| 当前分支与文件名 | 确认正在编辑的稿件和分支，点击“刷新”重新加载记录。 |
| 文档状态 | 查看当前稿是否有未保存或未提交的修改。“对比最新提交”比较当前稿与当前分支最新提交，包括未保存的编辑。 |
| 远程协作 | 显示远程分支及领先、落后数量，提供“关联 / 获取 / 拉取 / 推送”四个按钮。 |
| 分支筛选 | 选择全部分支、本地分支或远程分支。这里改变提交图的显示范围；切换正文使用功能区“切换分支”。 |
| 本地操作按钮 | 在原稿侧栏可提交版本或处理返稿。合并稿侧栏显示剩余修订，处理完后可点击“完成合并”。 |
| 提交图 | 查看提交之间的连接及分支、标签标记。选中后预览或对比，双击直接预览；按住 Ctrl 选择两版可对比所选版本。 |
| 底部信息 | 查看版本与作者，打开项目链接或检查更新。 |

`origin/main` 表示已获取的远程 `main` 分支。“领先”是本地尚未推送的提交数，“落后”是远程尚未合入的提交数；数量以最近一次获取或推送为准。

## 常用流程

### 接收老师返稿

1. 提交当前稿，点击 **返稿合并**。选择返回的 DOCX 和实际发给老师的送审版本，核对合入分支。
2. 需要在提交图中保留原始返稿时，勾选 **将原始返稿保存为独立分支**。
3. 点击 **生成合并稿**，在右侧窗格查看批注和剩余修订。处理完修订后点击 **完成合并**。

也可以先点 **仅保存为分支**，之后通过 **合并分支** 继续。

### 通过 GitHub 协作

- **发起人**：提交本地稿件 → 在 GitHub 创建空仓库并给协作者开放访问权限 → 关联仓库 → 推送。
- **其他作者**：点击 **克隆稿件**，填写相同的仓库地址和分支，选择新的本地 DOCX 保存路径。
- **日常同步**：提交当前修改 → 获取 → 拉取 → 推送。

双方都有新增提交时，在生成的 Word 合并稿中逐条审阅，完成后点击 **提交版本**。回到原稿推送，其他作者即可拉取确认后的结果。

HTTPS 连接使用 Git for Windows 的凭据管理器，首次连接按提示登录；SSH 使用已配置的密钥，自己设置的 `core.sshCommand` 会被保留。远程仓库需包含 WordGit 历史，供协作者通过插件还原稿件。网络操作进行中可随时点击 **取消**，不会留下半个版本；长时间没有任何传输进展时会自动停止。远程历史被强制改写时，插件会保留改写前已下载的版本，防止本地版本丢失。

## 项目结构

```text
WordGit/
├─ src/                 Word 插件、侧栏、版本管理和远程协作
├─ packaging/           独立安装程序及其界面验证
├─ tests/               历史、远程同步、Word、界面和安装测试
├─ docs/
│  ├─ USAGE.md          完整使用流程
│  ├─ DEVELOPMENT.md    构建、测试和开发说明
│  └─ RELEASING.md      打包、发布与验证记录
├─ .github/             CI、CodeQL、PR 与 Issue 模板
├─ build.ps1            编译插件并运行测试
├─ package.ps1          生成 EXE、ZIP 和 SHA-256 校验文件
├─ install.ps1          当前用户安装、升级和卸载
├─ Install.cmd          安装入口
├─ Check.cmd            环境检查入口
├─ Uninstall.cmd        卸载入口
├─ bin/                 编译与测试输出（生成）
├─ dist/                发布包（生成）
├─ README.md
└─ LICENSE
```

`src/` 中的主要模块：

| 文件 | 职责 |
| --- | --- |
| `AddIn.cs` | 插件入口、WordGit 功能区和各项操作的流程。 |
| `Operations.cs` | 按稿件区分的后台操作状态、退出前等待，以及界面线程同步上下文。 |
| `TaskPanes.cs`、`HistoryPane.cs`、`PaneLoadingFrame.cs`、`RepositoryView.cs`、`CommitGraph.cs` | 右侧窗格、加载框架、批注与修订列表、提交图。 |
| `WordDocuments.cs` | Word 文档操作、原生对比与修订合并、替换当前稿。 |
| `WorkingCopyJournal.cs` | 替换当前稿的意图记录；Word 中途退出后启动时自动还原或完成。 |
| `GitHistory.cs` | 本地提交、分支、历史查询、版本取回和存储维护。 |
| `GitProcess.cs` | 调用 git.exe：环境隔离、落盘同步、超时与整棵进程树的停止。 |
| `RemoteSync.cs`、`RemoteDialog.cs` | 远程连接、克隆、获取、拉取和推送。 |
| `ImportReviewDialog.cs` | 返稿文件选择、原始返稿分支和合并选项。 |
| `DocumentSnapshot.cs`、`DocxFormat.cs` | 文档分块快照、DOCX 校验与旧历史格式读取。 |
| `HistoryMaintenance.cs` | 空闲时自动整理仓库存储。 |
| `Dialogs.cs`、`DialogPlacement.cs` | 提交标题和详细内容、分支和历史操作对话框、居中显示的进度窗口。 |
| `AppInfo.cs`、`AboutDialog.cs`、`UpdateChecker.cs` | 版本信息、作者链接和更新检查。 |

## 数据位置与备份

| 目录 | 内容 |
| --- | --- |
| `%LOCALAPPDATA%\WordGit\app` | 插件和安装维护程序。 |
| `%LOCALAPPDATA%\WordGit\repositories` | 稿件的 Git 历史仓库。 |
| `%LOCALAPPDATA%\WordGit\sessions` | 预览、对比和合并审阅文件；一天前的残留会自动清理。 |
| `%LOCALAPPDATA%\WordGit\logs` | 出错记录 `wordgit.log`（超过 1 MB 自动轮转），反馈问题时请附上。日志中的用户目录写作 `%USERPROFILE%`。 |

历史按稿件的完整路径关联。稿件改名或移动后会对应新的历史位置，原历史仓库仍保留。

切换分支、恢复旧版或拉取时，原稿会先备份到同一文件夹中的 `.wordgit-restore-*.tmp.bak`，成功后立即删除。如果 Word 在替换途中退出，下次启动时插件会自动还原原稿或完成操作；无法自动处理时会提示备份位置，这类备份保留 30 天。后台提交、推送等操作尚未完成时关闭 Word，Word 会等待它们结束。

备份时关闭 Word，同时保存原稿和整个 `repositories` 文件夹；有未完成的合并审阅时，再保存 `sessions`。仓库存储由插件在后台自动整理，日常使用无需手动维护。

## 许可证

Copyright (C) 2026 QTH1225。

本项目采用 GNU General Public License v3.0（SPDX：`GPL-3.0-only`）。使用、修改和再分发遵循 [LICENSE](LICENSE)。分发安装包时，同时提供对应版本的完整源码和构建脚本。

## 开发与更多说明

在项目根目录运行：

```powershell
# 构建并检查功能与界面
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Test -UiTest

# 生成并验证发布包
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\package.ps1 -Test
```

- [使用说明](docs/USAGE.md)：分支合并、返稿导入、远程协作和版本取回的详细步骤。
- [开发说明](docs/DEVELOPMENT.md)：源码安装、测试命令和模块开发。
- [发布说明](docs/RELEASING.md)：安装包制作、发布检查及兼容性记录。
