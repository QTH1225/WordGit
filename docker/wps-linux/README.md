# Linux WPS 探测环境

在 Docker 中运行 Linux 版 WPS 个人版、WordGit 探测加载项和探测服务，自动执行 `docs/RUST_PLAN.md` 中 P0 的检查项。

## 准备 WPS 安装包

安装包来自 WPS 官方 CDN，放到 `docker/wps-linux/cache/`（已加入 `.gitignore`）：

| 文件 | SHA-256 |
| --- | --- |
| `wps-office_12.1.2.29183_amd64.deb` | `77ffbdc1be76c834e1853c27ff9063e49d5d3b1f30b206d785761b5bf560c83b` |

下载地址需要带时间戳签名，生成方法见 AUR `wps-office-cn` 的 PKGBUILD。构建时会校验 SHA-256。

## 运行

在仓库根目录执行：

```powershell
docker compose -f docker/wps-linux/compose.yaml build
# 离线模式（默认，推荐）：加载项复制到 jsaddons/WordGitProbe_0.1.0，publish.xml 写法与 Zotero-WPSJS 相同
docker compose -f docker/wps-linux/compose.yaml up -d
# 在线模式：加载项页面由探测服务提供（jspluginonline）
$env:PROBE_MODE = "online"; docker compose -f docker/wps-linux/compose.yaml up -d
```

浏览器打开 <http://127.0.0.1:6080/vnc.html> 可观看 WPS 界面。

容器启动后会用 `probe-run.sh` 跑一遍探测：启动 WPS、自动确认"设为默认办公软件"和"是否信任加载项"两类对话框（只在探测开始前确认），然后等待探测完成或 WPS 崩溃，结果追加到 `runs.log`。之后可以重复运行，并切换窗口模式：

```powershell
docker exec -e WPS_WINDOW_MODE=independent wordgit-p0-wps-1 /opt/wordgit/probe-run.sh /out/report-offline.jsonl
```

| 变量 | 取值 |
| --- | --- |
| `WPS_WINDOW_MODE` | `fusion`（默认，整合窗口）或 `independent`（独立的 WPS 文字窗口） |

工作目录 `~/probe/probe-options.txt` 可以调整探测行为，每行一项：`skip <步骤名>` 跳过某步，`pause <毫秒>` 在文档操作之间额外等待。

镜像额外安装了 WPS 内置 CEF 需要、但 WPS 安装包没有声明的系统库（NSS/NSPR、ATK、AT-SPI、ALSA 等）。缺少它们时，加载项能被信任，却永远不会执行。

## 输出

`docker/wps-linux/out/` 下：

| 文件 | 内容 |
| --- | --- |
| `report-<模式>.jsonl` | 每个检查项一行：步骤名、成功与否、返回值或错误、请求的 Origin |
| `runs.log`、`dialogs.log` | 每轮探测的结果（`done`/`crashed`/`timeout`）和自动确认过的对话框 |
| `screen-<窗口模式>-<结果>.png` | 每轮结束时的截图 |
| `server-<模式>.log` | 探测服务收到的每个请求（Host、Origin、Referer、User-Agent） |
| `probe-summary-<模式>.txt` | 加载项通过 WPS"另存为文本"写出的完整结果，网络不通时以它为准 |
| `screen.png`、`shots/` | 当前截图和每 15 秒一张的历史截图 |
| `wps-<模式>.log`、`publish-<模式>.xml` | WPS 输出和实际使用的加载项登记文件 |
