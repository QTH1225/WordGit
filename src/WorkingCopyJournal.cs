using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace WordGit
{
    // Makes "replace the working copy, then publish refs" recoverable. Before the file is replaced,
    // an intent record is written into the repository; it is removed once both steps are done. If
    // Word dies in between, Recover() finds the record and either rolls the file back (refs were
    // not published) or keeps it (they were), so the document and HEAD never silently disagree.
    internal static class WorkingCopyJournal
    {
        internal const string FileName = "wordgit-replace.journal";
        private const string Format = "WORDGIT-REPLACE 1";
        private static readonly Encoding Utf8 = new UTF8Encoding(false, true);

        internal sealed class Entry
        {
            public string Document, Backup, OriginalHash, ReplacementHash, References;
        }

        internal static string PathFor(GitHistory repo) { return Path.Combine(repo.RepositoryPath, FileName); }

        internal static void Write(GitHistory repo, Entry entry)
        {
            var text = String.Join("\n", Format, entry.Document, entry.Backup, entry.OriginalHash, entry.ReplacementHash,
                Convert.ToBase64String(Utf8.GetBytes(entry.References))) + "\n";
            var bytes = Utf8.GetBytes(text);
            var path = PathFor(repo);
            var temporary = path + ".tmp";
            using (var stream = new FileStream(temporary, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                stream.Write(bytes, 0, bytes.Length);
                stream.Flush(true);
            }
            if (File.Exists(path)) File.Delete(path);
            File.Move(temporary, path);
        }

        internal static void Clear(GitHistory repo)
        {
            var path = PathFor(repo);
            if (File.Exists(path)) File.Delete(path);
        }

        internal static Entry Read(GitHistory repo)
        {
            var path = PathFor(repo);
            if (!File.Exists(path)) return null;
            var lines = File.ReadAllText(path, Utf8).TrimEnd('\n').Split('\n');
            if (lines.Length != 6 || lines[0] != Format || !Path.IsPathRooted(lines[1]) || !Path.IsPathRooted(lines[2]))
                throw new InvalidDataException("替换记录已损坏：" + path);
            return new Entry { Document = lines[1], Backup = lines[2], OriginalHash = lines[3], ReplacementHash = lines[4],
                References = Utf8.GetString(Convert.FromBase64String(lines[5])) };
        }

        // Returns a message for the user when something was repaired or needs attention.
        internal static string Recover(GitHistory repo)
        {
            Entry entry;
            try { entry = Read(repo); }
            catch (Exception error)
            {
                var path = PathFor(repo);
                File.Move(path, path + "." + DateTime.UtcNow.ToString("yyyyMMddHHmmss") + ".corrupt");
                return error.Message;
            }
            if (entry == null) return null;
            var current = Checksum(entry.Document);
            var published = repo.ReferenceState() != entry.References;
            string message = null;
            if (!published && current == entry.ReplacementHash && Checksum(entry.Backup) == entry.OriginalHash)
            {
                // The file was swapped but the branch never moved: put the original back.
                try { File.Replace(entry.Backup, entry.Document, null); }
                catch (IOException)
                {
                    // Usually the document is open in Word; the record stays for the next attempt.
                    return "上次切换或恢复在完成前中断，「" + Path.GetFileName(entry.Document) + "」需要还原为操作前的内容。" +
                        "请先在 Word 中关闭这份文档，再重试；原稿备份在：" + entry.Backup;
                }
                message = "上次切换或恢复在完成前中断，已将「" + Path.GetFileName(entry.Document) + "」还原为操作前的内容。";
            }
            else if (current == (published ? entry.ReplacementHash : entry.OriginalHash))
            {
                // Consistent already: either fully published, or interrupted before the swap.
                TryDelete(entry.Backup);
            }
            else if (File.Exists(entry.Backup))
            {
                message = "上次替换「" + Path.GetFileName(entry.Document) + "」时被中断，且文件之后又有变化。" +
                    "操作前的原稿备份保留在：" + entry.Backup;
            }
            Clear(repo);
            return message;
        }

        // Called when Word starts, before documents are opened, for every repository with a record.
        internal static void RecoverAll(string repositories, Action<string> report)
        {
            if (!Directory.Exists(repositories)) return;
            foreach (var record in Directory.EnumerateDirectories(repositories, "*.git")
                .Select(folder => Path.Combine(folder, FileName)).Where(File.Exists))
            {
                try
                {
                    var message = Recover(GitHistory.OpenRepository(Path.GetDirectoryName(record)));
                    if (message != null) report(message);
                }
                catch (Exception error) { AppPaths.LogError("working-copy recovery: " + error); }
            }
        }

        private static string Checksum(string path)
        {
            if (!File.Exists(path)) return null;
            using (var sha = SHA256.Create())
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                return GitHistory.Hex(sha.ComputeHash(stream));
        }

        private static void TryDelete(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
