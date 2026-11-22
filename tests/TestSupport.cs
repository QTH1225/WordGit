using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace WordGit
{
    // Shared by every test program. These helpers keep test code byte[]-friendly while routing
    // everything through the same file- and stream-based APIs that the add-in uses.
    internal static class TestRepository
    {
        public static byte[] Extract(this GitHistory repo, string revision)
        {
            using (var output = new MemoryStream())
            {
                repo.ExtractTo(revision, output);
                return output.ToArray();
            }
        }

        public static string Commit(this GitHistory repo, byte[] bytes, string message, string author, string expectedHead,
            string branch = null, string sourceBranch = null, string sourceTip = null)
        {
            var path = TestFiles.Temporary(bytes);
            try { return repo.CommitFile(path, message, author, expectedHead, branch, sourceBranch, sourceTip); }
            finally { File.Delete(path); }
        }
    }

    internal static class TestWord
    {
        public static object Review(this WordDocuments word, byte[] original, byte[] revised, bool merge, string author = null)
        {
            var left = TestFiles.Temporary(original);
            var right = TestFiles.Temporary(revised);
            try { return word.ReviewPaths(left, right, merge, author); }
            finally { File.Delete(left); File.Delete(right); }
        }

        public static void ReplaceWorkingCopy(this WordDocuments word, GitHistory repo, string path, byte[] current, byte[] replacement, Action publish)
        {
            var staged = TestFiles.Temporary(replacement);
            try { word.ReplaceWorkingCopy(repo, path, TestFiles.Sha(current), staged, TestFiles.Sha(replacement), publish); }
            finally { File.Delete(staged); }
        }
    }

    internal static class TestFiles
    {
        public static string Temporary(byte[] bytes)
        {
            var path = Path.Combine(Path.GetTempPath(), "wordgit-test-" + Guid.NewGuid().ToString("N") + ".docx");
            File.WriteAllBytes(path, bytes);
            return path;
        }

        public static void WriteNew(string path, byte[] bytes)
        {
            using (var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                output.Write(bytes, 0, bytes.Length);
        }

        public static byte[] ReadDocx(string path)
        {
            DocxFormat.ValidateFile(path);
            using (var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var buffer = new MemoryStream())
            {
                input.CopyTo(buffer);
                return buffer.ToArray();
            }
        }

        public static byte[] Snapshot(object document) { return ReadDocx(WordDocuments.SaveDocument(document)); }

        public static string Sha(byte[] bytes) { return GitHistory.Sha(bytes); }
    }

    internal sealed class DocxPart
    {
        public string Name;
        public byte[] Bytes;
    }

    // Writer side of the retired "wordgit.meta" format, used to build compatibility fixtures.
    internal static class LegacyDocx
    {
        public static List<DocxPart> Read(byte[] bytes)
        {
            using (var archive = new ZipArchive(new MemoryStream(bytes, false), ZipArchiveMode.Read, false))
                return archive.Entries.Where(entry => !entry.FullName.EndsWith("/", StringComparison.Ordinal)).Select(entry =>
                {
                    using (var stream = entry.Open())
                    using (var buffer = new MemoryStream())
                    {
                        stream.CopyTo(buffer);
                        return new DocxPart { Name = entry.FullName, Bytes = buffer.ToArray() };
                    }
                }).ToList();
        }

        public static IEnumerable<DocxMetaEntry> Entries(IEnumerable<DocxPart> parts)
        {
            return parts.Select(part => new DocxMetaEntry { Name = part.Name, Size = part.Bytes.LongLength, Sha = TestFiles.Sha(part.Bytes) });
        }

        public static byte[] Fingerprint(IEnumerable<DocxPart> parts) { return DocxFormat.Fingerprint(Entries(parts)); }

        public static byte[] Fingerprint(byte[] docx) { return Fingerprint(Read(docx)); }

        public static string MetaText(IList<DocxPart> parts)
        {
            var builder = new StringBuilder("wordgit 2\nfingerprint " + DocxFormat.Hex(Fingerprint(parts)) + "\n");
            foreach (var part in parts)
                builder.Append(Convert.ToBase64String(Encoding.UTF8.GetBytes(part.Name))).Append(' ')
                    .Append(part.Bytes.LongLength).Append(' ').Append(TestFiles.Sha(part.Bytes)).Append('\n');
            return builder.ToString();
        }
    }
}
