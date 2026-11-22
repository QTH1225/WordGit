using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace WordGit
{
    internal sealed class DocxMetaEntry
    {
        public string Name, Sha;
        public long Size;
    }

    internal sealed class DocxMeta
    {
        public string Fingerprint;
        public List<DocxMetaEntry> Entries;
    }

    // DOCX container checks, plus the reader for the retired "wordgit.meta" storage format in which
    // each ZIP member was stored as its own blob. New commits use DocumentSnapshot chunks.
    internal static class DocxFormat
    {
        internal const string MetaName = "wordgit.meta";
        internal static readonly DateTimeOffset ZipTime = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
        private const string Magic = "wordgit 2";
        private const string FingerprintPrefix = "fingerprint ";
        private static readonly Encoding Utf8 = new UTF8Encoding(false, true);

        // Word already compresses these; the legacy format stored them uncompressed.
        private static readonly HashSet<string> StoredExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf",
            ".bin", ".odttf", ".ttf", ".otf", ".woff", ".woff2", ".mp3", ".mp4", ".wav", ".zip"
        };

        // Password-protected OOXML files are OLE compound files instead of zip archives.
        internal static bool IsEncrypted(byte[] header)
        {
            return header != null && header.Length >= 8 && header[0] == 0xD0 && header[1] == 0xCF &&
                header[2] == 0x11 && header[3] == 0xE0 && header[4] == 0xA1 && header[5] == 0xB1 &&
                header[6] == 0x1A && header[7] == 0xE1;
        }

        // Cheap structural check (signature and central directory only); used before committing
        // and before handing a restored file to Word.
        internal static void ValidateFile(string path)
        {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                var header = new byte[8];
                var read = 0;
                while (read < header.Length)
                {
                    var count = stream.Read(header, read, header.Length - read);
                    if (count <= 0) break;
                    read += count;
                }
                if (read == header.Length && IsEncrypted(header))
                    throw new InvalidDataException("文档已加密或受密码保护，请先在 Word 中取消密码保护后再提交。");
                stream.Position = 0;
                try
                {
                    using (var zip = new ZipArchive(stream, ZipArchiveMode.Read, true))
                        if (zip.GetEntry("[Content_Types].xml") == null || zip.GetEntry("word/document.xml") == null)
                            throw new InvalidDataException("文件不是可读取的 Word .docx。");
                }
                catch (InvalidDataException) { throw new InvalidDataException("文件不是可读取的 Word .docx。"); }
            }
        }

        internal static byte[] Fingerprint(IEnumerable<DocxMetaEntry> entries)
        {
            var ordered = entries.OrderBy(entry => entry.Name, StringComparer.Ordinal).ToArray();
            using (var outer = SHA256.Create())
            {
                var header = Utf8.GetBytes("wordgit-docx/2\n");
                outer.TransformBlock(header, 0, header.Length, null, 0);
                var zero = new[] { (byte)0 };
                var length = new byte[8];
                foreach (var entry in ordered)
                {
                    var name = Utf8.GetBytes(entry.Name);
                    outer.TransformBlock(name, 0, name.Length, null, 0);
                    outer.TransformBlock(zero, 0, 1, null, 0);
                    for (var index = 0; index < 8; index++) length[index] = (byte)(entry.Size >> (56 - index * 8));
                    outer.TransformBlock(length, 0, 8, null, 0);
                    var digest = FromHex(entry.Sha);
                    outer.TransformBlock(digest, 0, digest.Length, null, 0);
                }
                outer.TransformFinalBlock(new byte[0], 0, 0);
                return outer.Hash;
            }
        }

        internal static string Hex(byte[] bytes)
        {
            var builder = new StringBuilder(bytes.Length * 2);
            foreach (var value in bytes) builder.Append(value.ToString("x2", CultureInfo.InvariantCulture));
            return builder.ToString();
        }

        private static byte[] FromHex(string text)
        {
            var bytes = new byte[text.Length / 2];
            for (var index = 0; index < bytes.Length; index++)
                bytes[index] = Byte.Parse(text.Substring(index * 2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            return bytes;
        }

        internal static DocxMeta ParseMeta(string text)
        {
            var lines = (text ?? "").Split('\n');
            if (lines.Length < 2 || lines[0] != Magic || !lines[1].StartsWith(FingerprintPrefix, StringComparison.Ordinal))
                throw new InvalidDataException("无法读取此仓库的版本元数据。");
            var fingerprint = lines[1].Substring(FingerprintPrefix.Length).Trim();
            if (!IsHex64(fingerprint)) throw new InvalidDataException("无法读取此仓库的版本元数据。");
            var meta = new DocxMeta { Fingerprint = fingerprint, Entries = new List<DocxMetaEntry>() };
            for (var index = 2; index < lines.Length; index++)
            {
                if (lines[index].Length == 0) continue;
                var fields = lines[index].Split(' ');
                if (fields.Length != 3) throw new InvalidDataException("无法读取此仓库的版本元数据。");
                string name;
                long size;
                try { name = Utf8.GetString(Convert.FromBase64String(fields[0])); }
                catch (FormatException) { throw new InvalidDataException("无法读取此仓库的版本元数据。"); }
                catch (DecoderFallbackException) { throw new InvalidDataException("无法读取此仓库的版本元数据。"); }
                ValidateName(name);
                if (!Int64.TryParse(fields[1], NumberStyles.None, CultureInfo.InvariantCulture, out size) || size < 0)
                    throw new InvalidDataException("无法读取此仓库的版本元数据。");
                if (!IsHex64(fields[2])) throw new InvalidDataException("无法读取此仓库的版本元数据。");
                meta.Entries.Add(new DocxMetaEntry { Name = name, Sha = fields[2], Size = size });
            }
            return meta;
        }

        private static bool IsHex64(string value)
        {
            return value.Length == 64 && value.All(character => "0123456789abcdef".IndexOf(character) >= 0);
        }

        internal static void ValidateName(string name)
        {
            if (String.IsNullOrEmpty(name) || name.Length > 4096 || name[0] == '/' || name.IndexOf('\\') >= 0 ||
                name.IndexOf('"') >= 0)
                throw new InvalidDataException("Word 文档中含有无法保存的部件名。");
            if (String.Equals(name, MetaName, StringComparison.OrdinalIgnoreCase) ||
                name.StartsWith(MetaName + "/", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Word 文档包含与 WordGit 元数据同名的部件，暂时无法保存。");
            foreach (var character in name)
                if (character < 0x20 || character == 0x7f)
                    throw new InvalidDataException("Word 文档中含有无法保存的部件名。");
            foreach (var segment in name.Split('/'))
                if (segment.Length == 0 || segment == "." || segment == "..")
                    throw new InvalidDataException("Word 文档中含有无法保存的部件名。");
        }

        internal static CompressionLevel LevelFor(string name)
        {
            var extension = Path.GetExtension(name);
            return extension.Length > 0 && StoredExtensions.Contains(extension)
                ? CompressionLevel.NoCompression : CompressionLevel.Optimal;
        }
    }
}
