using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
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
            using (var stream = new MemoryStream(bytes, false))
                return CutChunks(stream, bytes.Length, null, null)
                    .Select(chunk => new ArraySegment<byte>(bytes, checked((int)chunk.Offset), chunk.Length)).ToList();
        }

        // Streaming counterpart of Split: reads the document sequentially and keeps only a bounded
        // per-chunk buffer, so planning a large document needs only a few megabytes of memory.
        public static SnapshotPlan Plan(Stream stream, string objectFormat)
        {
            if (objectFormat != "sha1" && objectFormat != "sha256") throw new InvalidOperationException("不支持此 Git 对象格式。");
            if (!stream.CanSeek) throw new ArgumentException("需要可定位的文档流。");
            var length = stream.Length;
            if (length < 4) throw new ArgumentException("无法读取文档数据。");
            using (var overall = SHA256.Create())
            using (var hash = objectFormat == "sha1" ? (HashAlgorithm)SHA1.Create() : SHA256.Create())
            {
                var objects = new List<string>();
                var chunks = CutChunks(stream, length, overall,
                    (content, size) => objects.Add(BlobId(hash, content, 0, size)));
                overall.TransformFinalBlock(new byte[0], 0, 0);
                var checksum = BitConverter.ToString(overall.Hash).Replace("-", "").ToLowerInvariant();
                long sum = 0;
                foreach (var chunk in chunks) sum += chunk.Length;
                if (sum != length) throw new InvalidDataException("文档分块长度校验失败。");
                return new SnapshotPlan(length, checksum,
                    chunks.Select(chunk => chunk.Offset).ToArray(),
                    chunks.Select(chunk => chunk.Length).ToArray(), objects.ToArray());
            }
        }

        private struct ChunkSlice { public long Offset; public int Length; }

        // Single chunking implementation shared by Split and Plan. Reads the stream segment by
        // segment in ascending order; onChunk (when given) receives each finished chunk's content
        // from a reusable buffer and overall (when given) hashes every byte in file order.
        private static List<ChunkSlice> CutChunks(Stream stream, long length, HashAlgorithm overall, Action<byte[], int> onChunk)
        {
            var boundaries = ZipBoundaries(stream, length);
            var chunks = new List<ChunkSlice>();
            var buffer = new byte[64 * 1024];
            var chunkBuffer = onChunk == null ? null : new byte[MaximumChunk];
            for (var part = 0; part < boundaries.Length - 1; part++)
            {
                var end = boundaries[part + 1];
                stream.Position = boundaries[part];
                var position = stream.Position;
                var chunkStart = position;
                var chunkLength = 0;
                ulong rolling = 0;
                while (position < end)
                {
                    var want = (int)Math.Min(buffer.Length, end - position);
                    ReadExact(stream, buffer, 0, want);
                    if (overall != null) overall.TransformBlock(buffer, 0, want, null, 0);
                    for (var i = 0; i < want; i++, position++)
                    {
                        rolling = unchecked((rolling << 1) + Gear[buffer[i]]);
                        if (chunkBuffer != null) chunkBuffer[chunkLength] = buffer[i];
                        chunkLength++;
                        if (chunkLength >= MinimumChunk && ((rolling & 0xFFFF) == 0 || chunkLength >= MaximumChunk))
                        {
                            chunks.Add(new ChunkSlice { Offset = chunkStart, Length = chunkLength });
                            if (onChunk != null) onChunk(chunkBuffer, chunkLength);
                            chunkStart = position + 1;
                            chunkLength = 0;
                            rolling = 0;
                        }
                    }
                }
                if (chunkStart < end)
                {
                    chunks.Add(new ChunkSlice { Offset = chunkStart, Length = chunkLength });
                    if (onChunk != null) onChunk(chunkBuffer, chunkLength);
                }
            }
            return chunks;
        }

        private static void ReadExact(Stream stream, byte[] buffer, int offset, int count)
        {
            while (count > 0)
            {
                var read = stream.Read(buffer, offset, count);
                if (read <= 0) throw new EndOfStreamException("文档在读取期间被修改。");
                offset += read;
                count -= read;
            }
        }

        private static long[] ZipBoundaries(Stream stream, long length)
        {
            // Keep each ZIP member's compressed payload separate from timestamps and directory records.
            // This reuses embedded pictures even when Word rewrites their surrounding ZIP metadata.
            var fallback = new long[] { 0, length };
            if (length < 22) return fallback;
            var boundaries = new SortedSet<long> { 0, length };
            try
            {
                var tailLength = (int)Math.Min(65537 + 20, length);
                var tail = new byte[tailLength];
                stream.Position = length - tailLength;
                ReadExact(stream, tail, 0, tailLength);
                var eocd = -1;
                for (var candidate = tailLength - 22; candidate >= 0; candidate--)
                    if (BitConverter.ToUInt32(tail, candidate) == 0x06054b50 &&
                        candidate + 22 + BitConverter.ToUInt16(tail, candidate + 20) == tailLength) { eocd = candidate; break; }
                if (eocd < 0) return fallback;
                var end = length - tailLength + eocd;
                var count = BitConverter.ToUInt16(tail, eocd + 10);
                var directory = (long)BitConverter.ToUInt32(tail, eocd + 16);
                if (directory > end || count == UInt16.MaxValue) return fallback;
                var directoryLength = end - directory;
                if (directoryLength > 4 * 1024 * 1024) return fallback; // Pathological central directory; chunk the file as-is.
                var records = new byte[(int)directoryLength];
                stream.Position = directory;
                ReadExact(stream, records, 0, records.Length);
                boundaries.Add(directory);
                var cursor = 0;
                var localHeader = new byte[30];
                for (var index = 0; index < count; index++)
                {
                    if (cursor > records.Length - 46 || BitConverter.ToUInt32(records, cursor) != 0x02014b50) return fallback;
                    var packedSize = (long)BitConverter.ToUInt32(records, cursor + 20);
                    var localOffset = (long)BitConverter.ToUInt32(records, cursor + 42);
                    if (localOffset > directory - 30 || packedSize > length) return fallback;
                    stream.Position = localOffset;
                    ReadExact(stream, localHeader, 0, localHeader.Length);
                    if (BitConverter.ToUInt32(localHeader, 0) != 0x04034b50) return fallback;
                    var payload = localOffset + 30 + BitConverter.ToUInt16(localHeader, 26) + BitConverter.ToUInt16(localHeader, 28);
                    var payloadEnd = payload + packedSize;
                    if (payload > directory || payloadEnd > directory) return fallback;
                    boundaries.Add(localOffset);
                    boundaries.Add(payload);
                    boundaries.Add(payloadEnd);
                    cursor += 46 + BitConverter.ToUInt16(records, cursor + 28) +
                        BitConverter.ToUInt16(records, cursor + 30) + BitConverter.ToUInt16(records, cursor + 32);
                }
                return boundaries.ToArray();
            }
            catch (ArgumentException) { return fallback; }
        }

        public static string[] BlobObjects(IList<ArraySegment<byte>> chunks, string objectFormat)
        {
            if (objectFormat != "sha1" && objectFormat != "sha256") throw new InvalidOperationException("不支持此 Git 对象格式。");
            var objects = new string[chunks.Count];
            using (var hash = objectFormat == "sha1" ? (HashAlgorithm)SHA1.Create() : SHA256.Create())
                for (var i = 0; i < chunks.Count; i++)
                    objects[i] = BlobId(hash, chunks[i].Array, chunks[i].Offset, chunks[i].Count);
            return objects;
        }

        private static string BlobId(HashAlgorithm hash, byte[] content, int offset, int count)
        {
            hash.Initialize();
            var header = Encoding.ASCII.GetBytes("blob " + count.ToString(CultureInfo.InvariantCulture) + "\0");
            hash.TransformBlock(header, 0, header.Length, header, 0);
            hash.TransformFinalBlock(content, offset, count);
            return BitConverter.ToString(hash.Hash).Replace("-", "").ToLowerInvariant();
        }

        public static byte[] Manifest(byte[] bytes, IList<ArraySegment<byte>> chunks, string[] objects)
        {
            return Manifest(bytes.Length, GitHistory.Sha(bytes),
                chunks.Select(chunk => chunk.Count).ToArray(), objects);
        }

        public static byte[] Manifest(long length, string checksum, int[] sizes, string[] objects)
        {
            if (sizes.Length != objects.Length) throw new InvalidDataException("分块写入数量不一致。");
            var lines = new List<string> { Format, length.ToString(CultureInfo.InvariantCulture), checksum };
            for (var i = 0; i < objects.Length; i++)
                lines.Add(objects[i] + " " + sizes[i].ToString(CultureInfo.InvariantCulture));
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

        // Streaming counterpart of Restore(byte[]): verifies and writes each chunk as it arrives,
        // so memory stays bounded no matter how large the restored document is.
        public void Restore(Stream batch, Stream output)
        {
            var buffer = new byte[64 * 1024];
            using (var sha = SHA256.Create())
            {
                for (var i = 0; i < Objects.Length; i++)
                {
                    var header = ReadBatchHeader(batch);
                    var expected = Objects[i] + " blob " + Sizes[i].ToString(CultureInfo.InvariantCulture);
                    if (header != expected) throw new InvalidDataException("历史数据块标识或长度校验失败。");
                    var remaining = Sizes[i];
                    while (remaining > 0)
                    {
                        var step = Math.Min(buffer.Length, remaining);
                        var read = 0;
                        while (read < step)
                        {
                            var count = batch.Read(buffer, read, step - read);
                            if (count <= 0) throw new InvalidDataException("历史数据块不完整。");
                            read += count;
                        }
                        sha.TransformBlock(buffer, 0, step, null, 0);
                        output.Write(buffer, 0, step);
                        remaining -= step;
                    }
                    if (batch.ReadByte() != '\n') throw new InvalidDataException("历史数据块不完整。");
                }
                if (batch.ReadByte() >= 0) throw new InvalidDataException("历史文档校验失败。");
                sha.TransformFinalBlock(new byte[0], 0, 0);
                if (BitConverter.ToString(sha.Hash).Replace("-", "").ToLowerInvariant() != Checksum)
                    throw new InvalidDataException("历史文档校验失败。");
            }
        }

        private static string ReadBatchHeader(Stream batch)
        {
            var bytes = new List<byte>(64);
            for (;;)
            {
                var next = batch.ReadByte();
                if (next < 0 || bytes.Count > 100) throw new InvalidDataException("无法读取历史数据块。");
                if (next == '\n') return Encoding.ASCII.GetString(bytes.ToArray());
                bytes.Add((byte)next);
            }
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

    // Metadata for a file whose chunks were planned without keeping the content in memory.
    internal sealed class SnapshotPlan
    {
        public readonly long Length;
        public readonly string Checksum;
        public readonly long[] Offsets;
        public readonly int[] Sizes;
        public readonly string[] Objects;

        internal SnapshotPlan(long length, string checksum, long[] offsets, int[] sizes, string[] objects)
        {
            Length = length;
            Checksum = checksum;
            Offsets = offsets;
            Sizes = sizes;
            Objects = objects;
        }
    }
}
