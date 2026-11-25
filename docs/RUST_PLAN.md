# WordGit Rust 内核与 WPS 支持方案

分支：`rust-core`。目标：让 WordGit 在 Windows、Linux、macOS 的 WPS 文字中可用，同时保留 Windows 桌面版 Word 的现有能力。

## 结论

兼容 WPS 的关键是插件形态，内核语言是第二位的问题：

- 现有 WordGit 是 .NET COM 加载项，界面使用 WinForms。WPS 在 Linux 和 macOS 上不支持 COM。
- WPS 跨平台的扩展方式是 **JS 加载项**（wpsjs）：HTML + JS 运行在 WPS 内置浏览器中，`window.Application` 提供仿 VBA 的文字对象模型。
- JS 加载项无法启动本地进程（`OAAssist.ShellExecute` 仍可调用，但官方已以安全为由让它失效），也不能直接运行 Git。所以 Rust 服务要跟随用户登录自启，不能依赖加载项按需拉起。

因此采用 **Rust 本地服务 + Web 加载项** 的结构（已实现，见下文「当前状态」）：

```text
crates/wordgit-core     分块快照、DOCX 校验、提交/分支/历史、远程同步、恢复记录、存储整理
crates/wordgit          二进制 wordgit：本地 HTTP 服务（serve）、安装（install）与命令行
addin-wps/              WPS JS 加载项：功能区、"批注、修订与提交"侧栏、对话框
src/ (C#)               Windows Word COM 加载项，与 Rust 内核共用同一份仓库格式
```

同一套 Web 界面和服务以后也可用于 Word for Mac（Office.js）。

## 当前状态（2026-11-25）

| 部分 | 状态 | 验证 |
| --- | --- | --- |
| `wordgit-core` | 全部移植：三种存储格式读写、分块快照、提交/分支/切换/删除、返稿导入、合并提交、远程关联/获取/推送/快进/克隆、强推备份、替换当前稿与崩溃恢复、后台压缩 | 22 个单元测试、8 个集成测试（真实 git，含双客户端远程），Windows 与 Linux 均通过 |
| 与 C# 版互通 | 仓库位置、分块快照树、整文件校验和完全一致；两边互相读取对方的提交并逐字节还原 | `tests/rust-compat.ps1`（8 项） |
| `wordgit` 服务 | 25 个接口；令牌、Host、Origin 校验；长任务取消；替换计划（准备 → 编辑器关闭文档 → 应用） | 服务 HTTP 测试（安全检查、提交/分支/切换/恢复往返） |
| WPS 加载项 | 功能区 17 个按钮、侧栏（批注与修订、状态、远程、提交图）、9 个对话框 | Docker 内 Linux WPS 个人版端到端 20 步，整合窗口与独立窗口模式 |

### 合并在 WPS 中的做法

Word 的 `MergeDocuments` 能把两份各带修订的稿子合成一份三方合并稿；WPS 实测做不到：

| 调用 | WPS 12.1.2 实测 |
| --- | --- |
| `Application.MergeDocuments(a, b, …)` | 只是激活第二份文档，不生成合并稿 |
| `Document.Merge(file, wdMergeTargetCurrent)` | 把两份的差异写成修订，但会丢掉两侧原有修订的作者，所有修订作者都变成"未知" |
| `Application.CompareDocuments(a, b, …)` | 正常，修订作者可指定 |

所以 WordGit 在 WPS 上统一用 `CompareDocuments` 生成合并稿：

- **当前分支自共同版本以来没有修改**（例如老师返稿后自己还没动过、拉取时只有对方有新提交）：用"当前稿 → 对方版本"对比，结果正好是对方的修改，作者显示为对方，精确无误。
- **双方都有修改**：合并稿里除了对方的修改，还会出现"撤销当前分支修改"的修订。加载项会另外打开一份"当前分支的修改"对照稿，并提示作者把与对照稿相反的修订拒绝掉。这比 Word 版多一步，但不会丢内容。

## 已核实的事实

| 事项 | 结论 | 来源 |
| --- | --- | --- |
| 原生比较 | `Application.CompareDocuments`、`Document.Compare` 在 Linux 个人版 12.1.2 实测可用，能生成带修订和原批注的对比稿。参数顺序与 Word 完全一致，第 13 位是 `CompareMoves`（旧版 WPS 文档漏写了它）。JS 只能按位置传参 | P0 实测 |
| 合并 | `Application.MergeDocuments` 不生成合并稿；`Document.Merge` 丢失修订作者。改用 `CompareDocuments`，见上文「合并在 WPS 中的做法」 | P0 实测 |
| 启动本地进程 | `OAAssist.ShellExecute` 已失效。社区插件改用 JSA 宏的 `Shell()` 绕过，但这要求用户降低宏安全级别，WordGit 不采用 | WPS 官方社区、Zotero-WPSJS |
| 本地文件 | `wps.FileSystem` 提供 `ReadFile`、`WriteFile`、`readdirSync`、`stat`、`copyFileSync` 等方法；`Application.Env` 提供主目录、临时目录、AppData 等系统路径 | P0 实测 |
| 加载项目录 | Windows：`%APPDATA%\kingsoft\wps\jsaddons`；Linux：`~/.local/share/Kingsoft/wps/jsaddons`；macOS：`~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons` | wpsjs 2.2.3 源码、Zotero-WPSJS |
| 推荐部署 | 把文件放到 `jsaddons/<名称>_<版本>/`，并在 `publish.xml` 中写 `<jsplugin url="http://127.0.0.1:<端口>/" type="wps" enable="enable_dev" install="null" version="<版本>" name="<名称>"/>`。页面从本地目录加载，origin 为 `file://` | Zotero-WPSJS（Gitee 1300+ 星）三平台安装脚本；P0 实测 |
| 平台支持 | Zotero-WPSJS 用同一方式支持最新的 Windows、Linux、macOS 版 WPS，其中 macOS 要求 WPS ≥ 6.15 | Zotero-WPSJS README、tankwyn/WPS-Zotero issue #34、#39 |
| 宏安全设置 | 不需要改动。保持默认的"中"、不勾选"信任对于 WPS JS 项目的访问"，加载项也能完整运行 | P0 实测 |
| 运行依赖 | WPS 内置 CEF 依赖 NSS/NSPR、ATK、AT-SPI、ALSA 等系统库，但 WPS 的安装包没有声明它们。缺少这些库时，加载项能被信任，却永远不会执行，WPS 自己的网页（如登录页）也会显示"页面加载失败 -1002"。桌面发行版一般自带这些库，Linux 安装包仍应声明依赖或在安装时检查 | P0 实测 |

## P0 可行性验证

用 `poc/` 下的探测加载项和探测服务，在每个目标平台上自动跑一组检查项，结果写入报告：

| 检查项 | 回答的问题 |
| --- | --- |
| `addin.loaded`、`start` | 加载项能否被该版本 WPS 载入；页面的 origin |
| `net.*` | 加载项能否 `fetch`/XHR 访问 127.0.0.1，能否通过带自定义请求头的预检 |
| `doc.*` | 新建、另存 DOCX、只读打开、编辑保存后重新打开 |
| `compare.*` | 三种比较/合并入口能否生成修订稿，修订作者、类型、页码 |
| `revisions.operate` | 单条接受/拒绝、批注读取、修订模式、全部接受 |
| `events.*` | 文档打开、保存、关闭等事件能否监听 |
| `ui.*`、`taskpane.*`、`dialog.*` | 任务窗格与对话框能否打开；其中的页面能否访问服务 |
| `host.capabilities` | 文件系统接口；`ShellExecute` 是否仍然存在 |

平台与执行方式：

| 平台 | 方式 | 状态 |
| --- | --- | --- |
| Linux WPS 个人版 12.1.2 | `docker/wps-linux/`：容器内全自动运行，noVNC 可观看 | **通过** |
| Windows WPS | 本机安装 WPS 后部署 `poc/wps-addin` 与探测服务 | 待做 |
| macOS WPS（M 芯片、Intel） | 真机。Docker 不能运行 macOS，社区的 Docker-OSX 只能模拟 Intel，而且不符合苹果许可 | 需要真机 |

### Linux 个人版 12.1.2.29183 结果（Docker，2026-11-24）

| 检查项 | 结果 |
| --- | --- |
| 载入 | 离线目录部署可以载入；页面地址为 `file:///…/jsaddons/WordGitProbe_0.1.0/index.html`，内核是 Chrome 104 |
| 本地服务 | 主页面、任务窗格、对话框都能通过 `fetch`/XHR 访问 `http://127.0.0.1:47820`。请求的 `Origin` 为 `file://`，带自定义请求头的预检也能通过 |
| 文档 | `SaveAs2(path, 16)` 保存 DOCX、只读打开、编辑保存后重新打开都正常 |
| 比较 | `CompareDocuments`、`Document.Compare` 都生成了 8 条修订（插入/删除，作者 WordGit），原批注保留；修订的页码可以读取 |
| 修订操作 | 选中、单条接受、单条拒绝、读取批注、打开修订模式、修订插入、全部接受都正常 |
| 事件 | `DocumentNew/Open/BeforeSave/BeforeClose/AfterClose/Change`、`WindowActivate` 都能注册，也都实际触发 |
| 文件系统 | `wps.FileSystem.WriteFile`/`ReadFile` 写入并读回中文内容 |
| 窗口模式 | 整合窗口（`AppComponentMode=prome_fushion`，WPS 默认）和独立组件窗口（`prome_independ`）都通过，前提是遵守下面的"让出事件循环"规则 |

**整合窗口模式的稳定性规则：文档操作之间必须让出事件循环。** 在一次同步执行中连续开关多个文档（新建、打开、比较、关闭），整合窗口模式下 WPS 每次都会崩溃（3/3），独立窗口模式不受影响（3/3 通过）。在每组文档操作之间 `await` 一次 `setTimeout(0)` 后，整合窗口模式 5/5 通过，探测耗时也从约 30 秒降到 12–14 秒。原因是整合窗口模式下打开、关闭文档要创建或销毁标签页，这些都是异步完成的。WordGit 加载项的文档操作一律写成异步流程，每一步之间让出事件循环。

需要注意的用户体验：

- 加载项首次载入、以及每次内容变化后（例如升级），WPS 都会弹出确认框（"首次加载，是否信任"/"已经被修改，是否继续加载"）。
- WPS 每次启动都会弹"设为默认办公软件"，与加载项无关。
- `Application.UserName` 未设置时返回"未知"，对比稿的修订作者应显式传入。

排查过程中走过的弯路：

- 一开始容器缺少 CEF 依赖库，加载项能被信任却从不执行，我据此误判"Linux 个人版不支持 JS 加载项"。后来参照 Zotero-WPSJS 的部署方式，又发现 WPS 自己的登录页同样加载失败，才定位到缺库。`docker/wps-linux/Dockerfile` 已补齐这些依赖。
- 自动确认对话框的脚本一度在探测进行中也会按回车，可能替探测"回答"了它自己触发的对话框。现在 `probe-run.sh` 只在探测开始之前确认对话框。

P0 的判定：

- Linux 已通过。Windows 上加载项能载入、能访问本地服务、能生成修订稿，就继续 P1。
- macOS 按社区经验应可行（WPS ≥ 6.15），仍需真机确认；不通过时降级为"尽力支持"，另评估 Word for Mac（Office.js）。

## 后续阶段

1. ~~**P1 Rust 内核**~~：已完成，并与 C# 版做了交叉验证。
2. ~~**P2 服务与 CLI**~~：已完成。登录自启：Linux 用 `~/.config/autostart/wordgit.desktop`，macOS 用 LaunchAgent，Windows 用 HKCU Run。
3. ~~**P3 WPS 加载项界面**~~：已完成，Linux WPS 端到端通过。
4. **P4 安装包**：Linux deb/rpm（声明 git 与 CEF 依赖库，安装后执行 `wordgit install`）、Windows、macOS。
5. **P5 统一（已完成）**：Windows Word 的 C# 加载项改为调用 Rust 服务，删除重复的 C# 历史实现；`src/` 只剩界面、Word COM 与 `ServiceClient`/`ServiceModels` 两个服务适配层。Word 原生对比与修订合并、后台 Word 生成合并稿、多轮返稿去重（`merge/context`）都保留。
6. **待实机验证**：Windows 版 WPS、macOS 版 WPS（沙盒内的会话目录、`FileDialog`）。

## 内核移植约束

- **存储格式**：必须读取三种格式（整包 `document.docx`、`wordgit.meta`、`wordgit.snapshot`）。分块算法（Gear 表、16 KiB–256 KiB 块大小、ZIP 边界切分）要与 C# 版逐字节一致，这样两端提交的数据块才能互相去重。
- **仓库定位**：C# 版使用 `SHA(完整路径.ToUpperInvariant())`。Windows 上保持不变，以兼容已有历史；Linux 区分大小写，不能转大写。macOS 默认文件系统不区分大小写，按卷属性决定。
- **Git**：先继续调用系统 git，保留现有的环境隔离（清除 `GIT_*`、屏蔽全局与系统配置、`core.fsync=committed`），网络操作沿用用户的凭据、SSH 和代理。以后本地对象读写可评估改用 gitoxide。
- **进程管理**：Windows 用作业对象，Unix 用进程组；超时或取消时结束整棵进程树。
- **替换当前稿**：保留意图记录和备份。Windows 用 `ReplaceFileW`，Unix 用同目录临时文件加 `rename`，并对文件和目录执行 fsync。加载项负责先在 WPS 中关闭文档，替换完成后再重新打开。
- **本地服务安全**：只绑定 127.0.0.1；校验 Host（防 DNS 重绑定）；所有请求都要带令牌。
- **令牌下发**：加载项页面从 `jsaddons` 目录加载，origin 是 `file://`，和服务不同源；而任何本地 HTML 文件的 origin 也都是 `file://`，所以 Origin 校验挡不住本机网页。服务启动时生成随机令牌，写入只有当前用户可读的文件，加载项用 `wps.FileSystem.ReadFile` 读取后放进请求头。
- **macOS 沙盒**：Mac 版 WPS 运行在 App Sandbox 中，加载项目录位于 `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/` 下。服务生成、需要由 WPS 打开的文件（预览、对比输入、合并稿）应放在 WPS 能访问的位置，需要在真机上验证。
