using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace WordGit
{
    internal sealed class DocumentSnapshot
    {
        internal const string ManifestFile = "wordgit.snapshot";
        internal const string ChunksFolder = "wordgit.chunks";
        private const string Format = "WORDGIT-SNAPSHOT 1";
        private const int MinimumChunk = 16 * 1024;
        private const int MaximumChunk = 256 * 1024;
        private static readonly ulong[] Gear = CreateGear();
        private static readonly Encoding Utf8 = new UTF8Encoding(false, true);
        public readonly int Length;
        public readonly string Checksum;
        public readonly string[] Objects;
        public readonly int[] Sizes;

        private DocumentSnapshot(int length, string checksum, string[] objects, int[] sizes)
        {
            Length = length;
            Checksum = checksum;
            Objects = objects;
            Sizes = sizes;
        }

        public static List<ArraySegment<byte>> Split(byte[] bytes)
        {
            var chunks = new List<ArraySegment<byte>>();
            var boundaries = ZipBoundaries(bytes);
            for (var part = 0; part < boundaries.Length - 1; part++)
            {
                var start = boundaries[part];
                ulong rolling = 0;
                for (var i = start; i < boundaries[part + 1]; i++)
                {
                    rolling = unchecked((rolling << 1) + Gear[bytes[i]]);
                    var size = i - start + 1;
                    if (size < MinimumChunk || ((rolling & 0xFFFF) != 0 && size < MaximumChunk)) continue;
                    chunks.Add(new ArraySegment<byte>(bytes, start, size));
                    start = i + 1;
                    rolling = 0;
                }
                if (start < boundaries[part + 1])
                    chunks.Add(new ArraySegment<byte>(bytes, start, boundaries[part + 1] - start));
            }
            return chunks;
        }

        private static int[] ZipBoundaries(byte[] bytes)
        {
            // Keep each ZIP member's compressed payload separate from timestamps and directory records.
            // This reuses embedded pictures even when Word rewrites their surrounding ZIP metadata.
            var boundaries = new SortedSet<int> { 0, bytes.Length };
            try
            {
                var end = bytes.Length - 22;
                while (end >= Math.Max(0, bytes.Length - 65557) &&
                    (BitConverter.ToUInt32(bytes, end) != 0x06054b50 ||
                        end + 22 + BitConverter.ToUInt16(bytes, end + 20) != bytes.Length)) end--;
                if (end < Math.Max(0, bytes.Length - 65557)) return boundaries.ToArray();

                var count = BitConverter.ToUInt16(bytes, end + 10);
                var directory = (long)BitConverter.ToUInt32(bytes, end + 16);
                if (directory > end || count == UInt16.MaxValue) return boundaries.ToArray();
                var cursor = (int)directory;
                boundaries.Add(cursor);
                for (var index = 0; index < count; index++)
                {
                    if (cursor < 0 || cursor > end - 46 || BitConverter.ToUInt32(bytes, cursor) != 0x02014b50)
                        return new[] { 0, bytes.Length };
                    var packedSize = (long)BitConverter.ToUInt32(bytes, cursor + 20);
                    var localOffset = (long)BitConverter.ToUInt32(bytes, cursor + 42);
                    if (localOffset > directory - 30 || packedSize > bytes.Length)
                        return new[] { 0, bytes.Length };
                    var local = (int)localOffset;
                    if (BitConverter.ToUInt32(bytes, local) != 0x04034b50)
                        return new[] { 0, bytes.Length };
                    var payload = localOffset + 30 + BitConverter.ToUInt16(bytes, local + 26) + BitConverter.ToUInt16(bytes, local + 28);
                    var payloadEnd = payload + packedSize;
                    if (payload > directory || payloadEnd > directory) return new[] { 0, bytes.Length };
                    boundaries.Add(local);
                    boundaries.Add((int)payload);
                    boundaries.Add((int)payloadEnd);
                    cursor += 46 + BitConverter.ToUInt16(bytes, cursor + 28) +
                        BitConverter.ToUInt16(bytes, cursor + 30) + BitConverter.ToUInt16(bytes, cursor + 32);
                }
                return boundaries.ToArray();
            }
            catch (ArgumentException) { return new[] { 0, bytes.Length }; }
        }

        public static string[] BlobObjects(IList<ArraySegment<byte>> chunks, string objectFormat)
        {
            if (objectFormat != "sha1" && objectFormat != "sha256") throw new InvalidOperationException("不支持此 Git 对象格式。");
            var objects = new string[chunks.Count];
            using (var hash = objectFormat == "sha1" ? (System.Security.Cryptography.HashAlgorithm)System.Security.Cryptography.SHA1.Create()
                : System.Security.Cryptography.SHA256.Create())
            {
                for (var i = 0; i < chunks.Count; i++)
                {
                    hash.Initialize();
                    var header = Encoding.ASCII.GetBytes("blob " + chunks[i].Count.ToString(CultureInfo.InvariantCulture) + "\0");
                    hash.TransformBlock(header, 0, header.Length, header, 0);
                    hash.TransformFinalBlock(chunks[i].Array, chunks[i].Offset, chunks[i].Count);
                    objects[i] = BitConverter.ToString(hash.Hash).Replace("-", "").ToLowerInvariant();
                }
            }
            return objects;
        }

        public static byte[] Manifest(byte[] bytes, IList<ArraySegment<byte>> chunks, string[] objects)
        {
            if (chunks.Count != objects.Length) throw new InvalidDataException("分块写入数量不一致。");
            var lines = new List<string> { Format, bytes.Length.ToString(CultureInfo.InvariantCulture), GitHistory.Sha(bytes) };
            for (var i = 0; i < chunks.Count; i++)
                lines.Add(objects[i] + " " + chunks[i].Count.ToString(CultureInfo.InvariantCulture));
            return Utf8.GetBytes(String.Join("\n", lines) + "\n");
        }

        public static DocumentSnapshot Read(byte[] bytes)
        {
            var lines = Utf8.GetString(bytes).TrimEnd('\n').Split('\n');
            int length;
            if (lines.Length < 4 || lines[0] != Format ||
                !Int32.TryParse(lines[1], NumberStyles.None, CultureInfo.InvariantCulture, out length) || length < 4 ||
                !Regex.IsMatch(lines[2], @"\A[a-f0-9]{64}\z"))
                throw new InvalidDataException("文档分块清单无效。");
            var objects = new string[lines.Length - 3];
            var sizes = new int[objects.Length];
            long sum = 0;
            for (var i = 0; i < objects.Length; i++)
            {
                var parts = lines[i + 3].Split(' ');
                if (parts.Length != 2 || !Regex.IsMatch(parts[0], @"\A(?:[a-f0-9]{40}|[a-f0-9]{64})\z") ||
                    !Int32.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out sizes[i]) ||
                    sizes[i] <= 0 || sizes[i] > MaximumChunk)
                    throw new InvalidDataException("文档分块清单损坏。");
                objects[i] = parts[0];
                sum += sizes[i];
            }
            if (sum != length) throw new InvalidDataException("文档分块长度校验失败。");
            return new DocumentSnapshot(length, lines[2], objects, sizes);
        }

        public byte[] Restore(byte[] batch)
        {
            var result = new byte[Length];
            var position = 0;
            var written = 0;
            for (var i = 0; i < Objects.Length; i++)
            {
                var end = Array.IndexOf(batch, (byte)'\n', position);
                if (end < 0 || end - position > 100) throw new InvalidDataException("无法读取历史数据块。");
                var header = Encoding.ASCII.GetString(batch, position, end - position);
                var expected = Objects[i] + " blob " + Sizes[i].ToString(CultureInfo.InvariantCulture);
                if (header != expected) throw new InvalidDataException("历史数据块标识或长度校验失败。");
                position = end + 1;
                if (batch.Length - position <= Sizes[i] || batch[position + Sizes[i]] != (byte)'\n')
                    throw new InvalidDataException("历史数据块不完整。");
                Buffer.BlockCopy(batch, position, result, written, Sizes[i]);
                position += Sizes[i] + 1;
                written += Sizes[i];
            }
            if (position != batch.Length || GitHistory.Sha(result) != Checksum)
                throw new InvalidDataException("历史文档校验失败。");
            return result;
        }

        private static ulong[] CreateGear()
        {
            var gear = new ulong[256];
            ulong seed = 0x9E3779B97F4A7C15;
            for (var i = 0; i < gear.Length; i++)
            {
                seed ^= seed >> 12;
                seed ^= seed << 25;
                seed ^= seed >> 27;
                gear[i] = unchecked(seed * 2685821657736338717);
            }
            return gear;
        }
    }

    internal sealed class PreparedSnapshot
    {
        public readonly string Tree, Checksum;
        public PreparedSnapshot(string tree, string checksum) { Tree = tree; Checksum = checksum; }
    }
}
