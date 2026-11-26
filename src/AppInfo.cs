using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace WordGit
{
    internal static class AppInfo
    {
        internal const string Author = "QTH1225";
        internal const string ProjectUrl = "https://github.com/QTH1225/WordGit";
        internal const string AuthorUrl = "https://github.com/QTH1225";
        internal const string ReleasesUrl = ProjectUrl + "/releases";
        internal const string IssuesUrl = ProjectUrl + "/issues";
        internal const string LicenseUrl = ProjectUrl + "/blob/main/LICENSE";
        internal const string ReleasesApi = "https://api.github.com/repos/QTH1225/WordGit/releases?per_page=25";
        internal static Version Version { get { return typeof(AppInfo).Assembly.GetName().Version; } }
        internal static string DisplayVersion { get { return ((System.Reflection.AssemblyInformationalVersionAttribute)Attribute.GetCustomAttribute(typeof(AppInfo).Assembly, typeof(System.Reflection.AssemblyInformationalVersionAttribute))).InformationalVersion; } }

        internal static bool IsAllowed(string url)
        {
            return url == ProjectUrl || url == AuthorUrl || url == ReleasesUrl || url == IssuesUrl || url == LicenseUrl || IsReleaseDownload(url);
        }

        private static bool IsReleaseDownload(string url)
        {
            if (String.IsNullOrEmpty(url) || !url.StartsWith(ReleasesUrl + "/", StringComparison.Ordinal)) return false;
            var parts = url.Substring(ReleasesUrl.Length + 1).Split('/');
            var version = parts.Length > 1 ? ReleaseVersion.Parse(parts[1]) : null;
            if (version == null) return false;
            return (parts.Length == 2 && parts[0] == "tag") ||
                (parts.Length == 3 && parts[0] == "download" && parts[2] == "WordGit-" + version.Text + "-Setup.exe");
        }
        internal static void OpenLink(IWin32Window owner, string url)
        {
            if (!IsAllowed(url))
            {
                MessageBox.Show(owner, "无效的 WordGit 链接。", "WordGit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
            catch (Exception) { MessageBox.Show(owner, "无法打开浏览器，请手动访问：" + Environment.NewLine + url,
                "WordGit", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        }
    }
}
