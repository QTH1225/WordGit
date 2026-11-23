using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Win32;

namespace WordGitSetup
{
    internal sealed class PackageFile
    {
        public string Path { get; set; }
        public string Sha256 { get; set; }
        public string Resource { get; set; }
    }

    internal sealed class PackageManifest
    {
        public string Version { get; set; }
        public List<PackageFile> Files { get; set; }
    }

    internal sealed class ScriptResult
    {
        public int Code;
        public string Output;
    }

    internal sealed class Payload : IDisposable
    {
        public string Root;
        public PackageManifest Manifest;

        public static string Hash(string path)
        {
            using (var algorithm = SHA256.Create())
            using (var input = File.OpenRead(path))
                return BitConverter.ToString(algorithm.ComputeHash(input)).Replace("-", "").ToLowerInvariant();
        }

        public static Payload Extract()
        {
            var result = new Payload { Root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "WordGit-Setup-" + Guid.NewGuid().ToString("N")) };
            Directory.CreateDirectory(result.Root);
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                string json;
                using (var stream = assembly.GetManifestResourceStream("WordGit.Manifest"))
                using (var reader = new StreamReader(stream, Encoding.UTF8)) json = reader.ReadToEnd();
                result.Manifest = new JavaScriptSerializer().Deserialize<PackageManifest>(json);
                var prefix = System.IO.Path.GetFullPath(result.Root) + System.IO.Path.DirectorySeparatorChar;
                if (result.Manifest.Files == null || result.Manifest.Files.Count == 0) throw new InvalidDataException("安装清单为空。");
                var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var file in result.Manifest.Files)
                {
                    if (String.IsNullOrWhiteSpace(file.Path) || System.IO.Path.IsPathRooted(file.Path) || !names.Add(file.Path))
                        throw new InvalidDataException("安装清单包含无效或重复路径。");
                    var target = System.IO.Path.GetFullPath(System.IO.Path.Combine(result.Root, file.Path));
                    if (!target.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("安装文件路径超出临时目录。");
                    Directory.CreateDirectory(System.IO.Path.GetDirectoryName(target));
                    using (var input = assembly.GetManifestResourceStream(file.Resource))
                    {
                        if (input == null) throw new InvalidDataException("安装文件缺失：" + file.Path);
                        using (var output = new FileStream(target, FileMode.CreateNew, FileAccess.Write)) input.CopyTo(output);
                    }
                    if (!String.Equals(Hash(target), file.Sha256, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("文件校验失败：" + file.Path);
                }
                File.WriteAllText(System.IO.Path.Combine(result.Root, "manifest.json"), json, new UTF8Encoding(true));
                var identity = AssemblyName.GetAssemblyName(System.IO.Path.Combine(result.Root, "bin", "WordGit.dll"));
                if (identity.Name != "WordGit" || identity.Version.ToString() != result.Manifest.Version) throw new InvalidDataException("安装包版本不一致。");
                var script = File.ReadAllBytes(System.IO.Path.Combine(result.Root, "install.ps1"));
                if (script.Length < 3 || script[0] != 239 || script[1] != 187 || script[2] != 191)
                    throw new InvalidDataException("安装脚本编码不兼容 Windows PowerShell。");
                return result;
            }
            catch { result.Dispose(); throw; }
        }

        public ScriptResult Run(string mode)
        {
            var executable = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell", "v1.0", "powershell.exe");
            var script = System.IO.Path.Combine(Root, "install.ps1");
            var arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File " + Quote(script);
            if (!String.IsNullOrEmpty(mode)) arguments += " " + mode;
            arguments += " -SetupExecutable " + Quote(Assembly.GetExecutingAssembly().Location);
            var start = new ProcessStartInfo(executable, arguments) {
                UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardOutput = true, RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8, WorkingDirectory = Root
            };
            // Started from a PowerShell 7 terminal, Setup inherits a module path whose 7.x modules
            // Windows PowerShell 5.1 cannot load; without it, 5.1 uses its own defaults.
            start.EnvironmentVariables.Remove("PSModulePath");
            using (var process = Process.Start(start))
            {
                var output = process.StandardOutput.ReadToEndAsync();
                var error = process.StandardError.ReadToEndAsync();
                process.WaitForExit();
                Task.WaitAll(output, error);
                return new ScriptResult { Code = process.ExitCode, Output = output.Result + error.Result };
            }
        }

        internal static string Quote(string value)
        {
            var output = new StringBuilder("\"");
            var slashes = 0;
            foreach (var character in value)
            {
                if (character == '\\') { slashes++; continue; }
                if (character == '"') { output.Append('\\', slashes * 2 + 1); output.Append(character); slashes = 0; continue; }
                output.Append('\\', slashes); slashes = 0; output.Append(character);
            }
            output.Append('\\', slashes * 2);
            return output.Append('"').ToString();
        }

        public void Dispose()
        {
            if (String.IsNullOrEmpty(Root)) return;
            var target = System.IO.Path.GetFullPath(Root);
            var parent = System.IO.Path.GetFullPath(System.IO.Path.GetTempPath()).TrimEnd('\\') + "\\";
            if (target.StartsWith(parent, StringComparison.OrdinalIgnoreCase) && System.IO.Path.GetFileName(target).StartsWith("WordGit-Setup-", StringComparison.Ordinal))
                try { Directory.Delete(target, true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
        }
    }

    internal sealed class SetupForm : Form
    {
        private readonly Payload payload;
        private readonly TextBox log = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };
        private readonly Button install = new Button { Text = "安装 / 更新", AutoSize = true };
        private readonly Button uninstall = new Button { Text = "卸载", AutoSize = true };
        private readonly Button check = new Button { Text = "检查环境", AutoSize = true };
        private readonly Button close = new Button { Text = "关闭", AutoSize = true };
        private readonly ProgressBar progress = new ProgressBar { Dock = DockStyle.Fill, Height = 14, Visible = false, Style = ProgressBarStyle.Marquee };
        private bool busy, removed;

        public SetupForm(Payload payload, bool uninstallMode)
        {
            this.payload = payload;
            Text = "WordGit 安装与维护";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(610, 420);
            MinimumSize = new Size(560, 400);
            StartPosition = FormStartPosition.CenterScreen;
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 6, Padding = new Padding(18) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(layout);
            layout.Controls.Add(new Label { Text = "WordGit " + new Version(payload.Manifest.Version).ToString(3), Font = new Font(Font.FontFamily, 20, FontStyle.Bold), AutoSize = true }, 0, 0);
            layout.Controls.Add(new Label { Text = "在 Word 或 WPS 文字中提交版本、查看分支、对比并处理修订。\r\n安装到当前 Windows 用户。请先保存并退出 Word 和 WPS 文字。", AutoSize = true, Margin = new Padding(0, 10, 0, 10) }, 0, 1);
            var links = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 10) };
            AddLink(links, "下载 Git for Windows", "https://git-scm.com/install/windows");
            AddLink(links, "下载 .NET Framework 4.8", "https://dotnet.microsoft.com/download/dotnet-framework/net48");
            layout.Controls.Add(links, 0, 2);
            log.Text = uninstallMode ? "点击“卸载”移除插件，版本历史和审阅稿会保留。" : "正在检查运行环境和安装文件……";
            layout.Controls.Add(log, 0, 3);
            layout.Controls.Add(progress, 0, 4);
            var buttons = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 12, 0, 0), WrapContents = true };
            foreach (var button in new[] { install, uninstall, check, close }) { button.Padding = new Padding(8, 3, 8, 3); buttons.Controls.Add(button); }
            layout.Controls.Add(buttons, 0, 5);
            install.Click += (sender, args) => Run("");
            uninstall.Click += (sender, args) => Run("-Uninstall");
            check.Click += (sender, args) => Run("-CheckOnly");
            close.Click += (sender, args) => Close();
            FormClosing += (sender, args) => { if (busy) args.Cancel = true; };
            FormClosed += (sender, args) => { if (removed) Program.ScheduleCleanup(); };
            if (!uninstallMode) Shown += (sender, args) => Run("-CheckOnly");
        }

        private static void AddLink(FlowLayoutPanel parent, string title, string url)
        {
            var link = new LinkLabel { Text = title, AutoSize = true, Margin = new Padding(0, 0, 16, 0) };
            link.LinkClicked += (sender, args) => Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            parent.Controls.Add(link);
        }

        private async void Run(string mode)
        {
            if (busy) return;
            busy = true;
            foreach (var button in new[] { install, uninstall, check, close }) button.Enabled = false;
            progress.Visible = true;
            log.Text = mode == "-Uninstall" ? "正在卸载插件，保留版本历史……" : mode == "-CheckOnly" ? "正在检查运行环境……" : "正在安装 WordGit……";
            try
            {
                var result = await Task.Run(() => payload.Run(mode));
                removed = result.Code == 0 && mode == "-Uninstall" || removed && mode == "-CheckOnly";
                log.Text = result.Output.Trim();
                if (result.Code != 0) log.AppendText("\r\n\r\n操作未完成。请根据上面的提示处理后重试。");
            }
            catch (Exception error) { log.Text = "操作未完成：" + error.Message; }
            finally
            {
                busy = false;
                progress.Visible = false;
                foreach (var button in new[] { install, uninstall, check, close }) button.Enabled = true;
            }
        }
    }

    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            if (args.Length == 2 && args[0] == "/cleanup") return CleanupInstalledCopy(Int32.Parse(args[1]));
            try
            {
                using (var payload = Payload.Extract())
                {
                    if (args.Length == 1 && args[0] == "/verify") { Console.WriteLine("PASS: embedded payload, SHA-256, version and script encoding."); return 0; }
                    if (args.Length == 1 && (args[0] == "/check" || args[0] == "/plan"))
                    {
                        var result = payload.Run(args[0] == "/plan" ? "-Plan" : "-CheckOnly");
                        Console.Write(result.Output);
                        return result.Code;
                    }
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new SetupForm(payload, args.Length == 1 && args[0] == "/uninstall"));
                }
                return 0;
            }
            catch (Exception error)
            {
                if (args.Length > 0) Console.Error.WriteLine(error.Message);
                else MessageBox.Show(error.Message, "WordGit 安装未完成", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
        }

        internal static void ScheduleCleanup()
        {
            try
            {
                var folder = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "WordGit-Setup-" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(folder);
                var copy = System.IO.Path.Combine(folder, "cleanup.exe");
                File.Copy(Assembly.GetExecutingAssembly().Location, copy);
                Process.Start(new ProcessStartInfo(copy, "/cleanup " + Process.GetCurrentProcess().Id) {
                    UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden
                });
            }
            catch (IOException) { } catch (UnauthorizedAccessException) { }
        }

        private static int CleanupInstalledCopy(int parentId)
        {
            try
            {
                try { using (var parent = Process.GetProcessById(parentId)) if (!parent.WaitForExit(30000)) return 1; }
                catch (ArgumentException) { }
                using (var registered = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\WordGit"))
                    if (registered != null) return 0;
                var parentFolder = System.IO.Path.GetFullPath(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordGit"));
                var folder = System.IO.Path.GetFullPath(System.IO.Path.Combine(parentFolder, "app"));
                if (!folder.StartsWith(parentFolder + System.IO.Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)) return 1;
                var executable = System.IO.Path.Combine(folder, "WordGit.Setup.exe");
                if (File.Exists(executable) && Payload.Hash(executable) == Payload.Hash(Assembly.GetExecutingAssembly().Location)) File.Delete(executable);
                if (Directory.Exists(folder) && Directory.GetFileSystemEntries(folder).Length == 0) Directory.Delete(folder, false);
                return 0;
            }
            catch (IOException) { return 1; } catch (UnauthorizedAccessException) { return 1; }
        }
    }
}
