using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace WordGit
{
    internal static class TrackedReviewSnapshot
    {
        private const string Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string Word2010 = "http://schemas.microsoft.com/office/word/2010/wordml";

        // Word can coalesce old and new adjacent edits in the native merge result. The raw
        // revisions still distinguish a cumulative return from a separate reply to the same draft.
        internal static Dictionary<string, int> Read(string path)
        {
            var result = new Dictionary<string, int>(StringComparer.Ordinal);
            using (var stream = File.OpenRead(path))
            using (var package = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                foreach (var part in package.Entries)
                {
                    if (!IsStory(part.FullName)) continue;
                    var document = new XmlDocument { XmlResolver = null };
                    using (var content = part.Open())
                    using (var reader = XmlReader.Create(content, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit }))
                        document.Load(reader);
                    var ns = new XmlNamespaceManager(document.NameTable);
                    ns.AddNamespace("w", Word);
                    foreach (XmlElement change in document.SelectNodes("//*[@w:author and @w:date]", ns))
                    {
                        if (change.NamespaceURI != Word || !IsRevision(change.LocalName)) continue;
                        var paragraph = (XmlElement)change.SelectSingleNode("ancestor::w:p[1]", ns);
                        var paragraphId = paragraph == null ? "" : paragraph.GetAttribute("paraId", Word2010);
                        // Without a stable anchor, fall back to the original submission-based merge.
                        if (paragraphId.Length == 0) return new Dictionary<string, int>(StringComparer.Ordinal);
                        var signature = (XmlElement)change.CloneNode(true);
                        signature.RemoveAttribute("id", Word);
                        var key = part.FullName + "\n" + paragraphId + "\n" + signature.OuterXml;
                        int count;
                        result.TryGetValue(key, out count);
                        result[key] = count + 1;
                    }
                }
            }
            return result;
        }

        private static bool IsStory(string name)
        {
            return name == "word/document.xml" || name == "word/footnotes.xml" || name == "word/endnotes.xml" ||
                (name.EndsWith(".xml", StringComparison.Ordinal) &&
                (name.StartsWith("word/header", StringComparison.Ordinal) || name.StartsWith("word/footer", StringComparison.Ordinal)));
        }

        private static bool IsRevision(string name)
        {
            return name == "ins" || name == "del" || name == "moveFrom" || name == "moveTo" || name.EndsWith("Change", StringComparison.Ordinal);
        }
    }
}
