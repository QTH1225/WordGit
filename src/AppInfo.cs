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
        internal const string LatestReleaseApi = "https://api.github.com/repos/QTH1225/WordGit/releases/latest";
        internal static Version Version { get { return typeof(AppInfo).Assembly.GetName().Version; } }
        internal static string DisplayVersion { get { return Version.ToString(3); } }

        internal static void OpenLink(IWin32Window owner, string url)
        {
            if (url != ProjectUrl && url != AuthorUrl && url != ReleasesUrl && url != IssuesUrl)
                throw new ArgumentException("无效的 WordGit 链接。");
            try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
            catch (Exception) { MessageBox.Show(owner, "无法打开浏览器，请手动访问：" + Environment.NewLine + url,
                "WordGit", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        }
    }
}
