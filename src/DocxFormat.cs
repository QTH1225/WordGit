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
    internal sealed class DocxPart
    {
        public string Name;
        public byte[] Bytes;
    }

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

    /// <summary>
    /// Stores a DOCX as the files inside its ZIP container so Git can delta the XML,
    /// instead of storing the already-deflated container as one opaque blob.
    /// </summary>
    internal static class DocxFormat
    {
        internal const string MetaName = "wordgit.meta";
        private const string Magic = "wordgit 2";
        private const string FingerprintPrefix = "fingerprint ";
        private static readonly DateTimeOffset ZipTime = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
        private static readonly Encoding Utf8 = new UTF8Encoding(false, true);

        // Word already compresses these; storing them as raw bytes lets Git reuse them between revisions.
        private static readonly HashSet<string> StoredExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf",
            ".bin", ".odttf", ".ttf", ".otf", ".woff", ".woff2", ".mp3", ".mp4", ".wav", ".zip"
        };

        internal static bool IsEncrypted(byte[] bytes)
        {
            // Password-protected OOXML files are OLE compound files instead of zip archives.
            return bytes != null && bytes.Length >= 8 && bytes[0] == 0xD0 && bytes[1] == 0xCF &&
                bytes[2] == 0x11 && bytes[3] == 0xE0 && bytes[4] == 0xA1 && bytes[5] == 0xB1 &&
                bytes[6] == 0x1A && bytes[7] == 0xE1;
        }

        private static ZipArchive OpenArchive(byte[] bytes)
        {
            try { return new ZipArchive(new MemoryStream(bytes, false), ZipArchiveMode.Read, false); }
            catch (InvalidDataException) { throw new InvalidDataException("文件不是可读取的 Word .docx。"); }
        }

        internal static List<DocxPart> Read(byte[] bytes)
        {
            if (bytes == null || bytes.Length < 4) throw new InvalidDataException("文件不是可读取的 Word .docx。");
            if (IsEncrypted(bytes)) throw new InvalidDataException("文档已加密或受保护，请先在 Word 中取消密码保护。");
            var parts = new List<DocxPart>();
            var names = new HashSet<string>(StringComparer.Ordinal);
            try
            {
                using (var archive = OpenArchive(bytes))
                {
                    foreach (var entry in archive.Entries)
                    {
                        if (entry.FullName.EndsWith("/", StringComparison.Ordinal)) continue;
                        var name = entry.FullName;
                        ValidateName(name);
                        if (!names.Add(name)) throw new InvalidDataException("Word 文档包含重复的部件：" + name);
                        byte[] content;
                        using (var stream = entry.Open())
                        using (var buffer = new MemoryStream())
                        {
                            stream.CopyTo(buffer);
                            content = buffer.ToArray();
                        }
                        parts.Add(new DocxPart { Name = name, Bytes = content });
                    }
                }
            }
            catch (InvalidDataException) { throw; }
            catch (Exception) { throw new InvalidDataException("文件不是可读取的 Word .docx。"); }
            if (parts.Count == 0) throw new InvalidDataException("文件不是可读取的 Word .docx。");

            var ordered = parts.Select(part => part.Name).OrderBy(name => name, StringComparer.Ordinal).ToArray();
            for (var index = 1; index < ordered.Length; index++)
                if (ordered[index].StartsWith(ordered[index - 1] + "/", StringComparison.Ordinal))
                    throw new InvalidDataException("Word 文档包含冲突的部件名：" + ordered[index - 1] + " / " + ordered[index]);
            return parts;
        }

        internal static byte[] Write(IList<DocxPart> parts)
        {
            using (var buffer = new MemoryStream())
            {
                using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, true))
                {
                    foreach (var part in parts)
                    {
                        var entry = archive.CreateEntry(part.Name, LevelFor(part.Name));
                        entry.LastWriteTime = ZipTime;
                        using (var stream = entry.Open()) stream.Write(part.Bytes, 0, part.Bytes.Length);
                    }
                }
                return buffer.ToArray();
            }
        }

        internal static byte[] Fingerprint(byte[] bytes)
        {
            return Fingerprint(Read(bytes));
        }

        internal static byte[] Fingerprint(IList<DocxPart> parts)
        {
            var ordered = parts.OrderBy(part => part.Name, StringComparer.Ordinal).ToArray();
            using (var outer = SHA256.Create())
            using (var inner = SHA256.Create())
            {
                var header = Utf8.GetBytes("wordgit-docx/2\n");
                outer.TransformBlock(header, 0, header.Length, null, 0);
                var zero = new[] { (byte)0 };
                var length = new byte[8];
                foreach (var part in ordered)
                {
                    var name = Utf8.GetBytes(part.Name);
                    outer.TransformBlock(name, 0, name.Length, null, 0);
                    outer.TransformBlock(zero, 0, 1, null, 0);
                    var size = part.Bytes.LongLength;
                    for (var index = 0; index < 8; index++) length[index] = (byte)(size >> (56 - index * 8));
                    outer.TransformBlock(length, 0, 8, null, 0);
                    var digest = inner.ComputeHash(part.Bytes);
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

        internal static string MetaText(byte[] fingerprint, IList<DocxPart> parts)
        {
            var builder = new StringBuilder();
            builder.Append(Magic).Append('\n');
            builder.Append(FingerprintPrefix).Append(Hex(fingerprint)).Append('\n');
            for (var index = 0; index < parts.Count; index++)
            {
                builder.Append(Convert.ToBase64String(Utf8.GetBytes(parts[index].Name))).Append(' ')
                    .Append(parts[index].Bytes.LongLength.ToString(CultureInfo.InvariantCulture)).Append(' ')
                    .Append(Hex(PartHash(parts[index].Bytes))).Append('\n');
            }
            return builder.ToString();
        }

        internal static byte[] PartHash(byte[] bytes)
        {
            using (var sha = SHA256.Create()) return sha.ComputeHash(bytes);
        }

        internal static DocxMeta ParseMeta(string text)
        {
            var lines = (text ?? "").Split('\n');
            if (lines.Length < 2 || lines[0] != Magic || !lines[1].StartsWith(FingerprintPrefix, StringComparison.Ordinal))
                throw new InvalidDataException("无法读取此仓库的版本元数据。");
            var fingerprint = lines[1].Substring(FingerprintPrefix.Length).Trim();
            if (fingerprint.Length != 64 || fingerprint.Any(character => "0123456789abcdef".IndexOf(character) < 0))
                throw new InvalidDataException("无法读取此仓库的版本元数据。");
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
                ValidateName(name);
                if (!Int64.TryParse(fields[1], NumberStyles.None, CultureInfo.InvariantCulture, out size) || size < 0)
                    throw new InvalidDataException("无法读取此仓库的版本元数据。");
                if (fields[2].Length != 64 || fields[2].Any(character => "0123456789abcdef".IndexOf(character) < 0))
                    throw new InvalidDataException("无法读取此仓库的版本元数据。");
                meta.Entries.Add(new DocxMetaEntry { Name = name, Sha = fields[2], Size = size });
            }
            return meta;
        }

        private static void ValidateName(string name)
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

        private static CompressionLevel LevelFor(string name)
        {
            var extension = Path.GetExtension(name);
            return extension.Length > 0 && StoredExtensions.Contains(extension)
                ? CompressionLevel.NoCompression : CompressionLevel.Optimal;
        }
    }
}