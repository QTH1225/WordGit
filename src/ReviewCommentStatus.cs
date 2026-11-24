using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;

namespace WordGit
{
    internal static class ReviewCommentStatus
    {
        private const string Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string Word2010 = "http://schemas.microsoft.com/office/word/2010/wordml";
        private const string Word2012 = "http://schemas.microsoft.com/office/word/2012/wordml";

        // The caller supplies a disposable copy of the merge result. Only the comment status part is written.
        internal static int Restore(string path, IEnumerable<DocumentComment> resolved)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.ReadWrite))
            using (var package = new ZipArchive(stream, ZipArchiveMode.Update))
            {
                var commentsPart = package.GetEntry("word/comments.xml");
                var statusPart = package.GetEntry("word/commentsExtended.xml");
                if (commentsPart == null || statusPart == null) return 0;

                var comments = ReadXml(commentsPart);
                var status = ReadXml(statusPart);
                var namespaces = new XmlNamespaceManager(comments.NameTable);
                namespaces.AddNamespace("w", Word);
                namespaces.AddNamespace("w15", Word2012);
                var threads = status.SelectNodes("/w15:commentsEx/w15:commentEx", namespaces).Cast<XmlElement>()
                    .Where(node => !node.HasAttribute("paraIdParent", Word2012))
                    .ToLookup(node => node.GetAttribute("paraId", Word2012), StringComparer.OrdinalIgnoreCase);
                var entries = comments.SelectNodes("/w:comments/w:comment", namespaces).Cast<XmlElement>()
                    .Select(comment => new
                    {
                        Author = comment.GetAttribute("author", Word),
                        Date = comment.GetAttribute("date", Word),
                        Text = CommentText(comment, namespaces),
                        Paragraph = (XmlElement)comment.SelectSingleNode("(.//w:p)[last()]", namespaces)
                    }).ToArray();

                var restored = 0;
                var changed = false;
                foreach (var candidate in resolved)
                {
                    var matches = entries.Where(entry => entry.Author == candidate.Author &&
                        entry.Text == NormalizeText(candidate.Text)).ToArray();
                    if (matches.Length > 1)
                        matches = matches.Where(entry => MatchesDate(entry.Date, candidate.Date)).ToArray();
                    if (matches.Length != 1 || matches[0].Paragraph == null) continue;
                    var paragraphId = matches[0].Paragraph.GetAttribute("paraId", Word2010);
                    if (paragraphId.Length == 0) continue;
                    var found = threads[paragraphId].ToArray();
                    if (found.Length != 1) continue;

                    var thread = found[0];
                    var done = thread.GetAttribute("done", Word2012);
                    if (done != "1" && done != "true")
                    {
                        var attribute = status.CreateAttribute(thread.Prefix, "done", Word2012);
                        attribute.Value = "1";
                        thread.SetAttributeNode(attribute);
                        changed = true;
                    }
                    restored++;
                }

                if (changed)
                {
                    var name = statusPart.FullName;
                    statusPart.Delete();
                    using (var destination = package.CreateEntry(name, CompressionLevel.Optimal).Open())
                        status.Save(destination);
                }
                return restored;
            }
        }

        private static XmlDocument ReadXml(ZipArchiveEntry entry)
        {
            var document = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
            using (var stream = entry.Open())
            using (var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit }))
                document.Load(reader);
            return document;
        }

        private static string CommentText(XmlElement comment, XmlNamespaceManager namespaces)
        {
            var paragraphs = comment.SelectNodes(".//w:p", namespaces).Cast<XmlElement>();
            return NormalizeText(string.Join("\n", paragraphs.Select(paragraph => string.Concat(
                paragraph.SelectNodes(".//w:t | .//w:tab | .//w:br | .//w:cr", namespaces).Cast<XmlElement>()
                .Select(node => node.LocalName == "t" ? node.InnerText : node.LocalName == "tab" ? "\t" : "\n")))));
        }

        private static string NormalizeText(string text)
        {
            return text.Replace("\r\n", "\n").Replace('\r', '\n').Replace('\v', '\n').Trim();
        }

        private static bool MatchesDate(string text, DateTime value)
        {
            DateTime parsed;
            return DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out parsed) &&
                (parsed == value || (parsed.Kind == DateTimeKind.Utc && parsed.ToLocalTime() == value));
        }
    }
}
