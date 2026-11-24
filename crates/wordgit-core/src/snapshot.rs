//! Content-defined chunking of DOCX files.
//!
//! A snapshot is a manifest (`wordgit.snapshot`) plus one blob per distinct chunk under
//! `wordgit.chunks/`. The file is first split at ZIP member boundaries, so each member's
//! compressed payload is chunked separately from timestamps and directory records; this reuses
//! embedded pictures even when Word rewrites their surrounding ZIP metadata. Within each part a
//! Gear rolling hash cuts chunks of 16–256 KiB.
//!
//! This must stay byte-for-byte identical to the original C# implementation: the same file must
//! produce the same chunks, blob ids and manifest so both clients deduplicate against each other.

use crate::{Error, Result};
use sha1::Sha1;
use sha2::{Digest, Sha256};
use std::collections::BTreeSet;
use std::io::{Read, Seek, SeekFrom, Write};

pub const MANIFEST_FILE: &str = "wordgit.snapshot";
pub const CHUNKS_FOLDER: &str = "wordgit.chunks";
const FORMAT: &str = "WORDGIT-SNAPSHOT 1";
pub const MINIMUM_CHUNK: usize = 16 * 1024;
pub const MAXIMUM_CHUNK: usize = 256 * 1024;
const BUFFER: usize = 64 * 1024;

static GEAR: [u64; 256] = create_gear();

const fn create_gear() -> [u64; 256] {
    let mut gear = [0u64; 256];
    let mut seed: u64 = 0x9E37_79B9_7F4A_7C15;
    let mut i = 0;
    while i < 256 {
        seed ^= seed >> 12;
        seed ^= seed << 25;
        seed ^= seed >> 27;
        gear[i] = seed.wrapping_mul(2_685_821_657_736_338_717);
        i += 1;
    }
    gear
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ObjectFormat {
    Sha1,
    Sha256,
}

impl ObjectFormat {
    pub fn parse(name: &str) -> Result<Self> {
        match name {
            "sha1" => Ok(ObjectFormat::Sha1),
            "sha256" => Ok(ObjectFormat::Sha256),
            _ => Err(Error::invalid("不支持此 Git 对象格式。")),
        }
    }
}

/// Git blob id of `content`.
pub fn blob_id(format: ObjectFormat, content: &[u8]) -> String {
    let header = format!("blob {}\0", content.len());
    match format {
        ObjectFormat::Sha1 => {
            let mut hash = Sha1::new();
            hash.update(header.as_bytes());
            hash.update(content);
            crate::hex(&hash.finalize())
        }
        ObjectFormat::Sha256 => {
            let mut hash = Sha256::new();
            hash.update(header.as_bytes());
            hash.update(content);
            crate::hex(&hash.finalize())
        }
    }
}

/// Chunk layout of a file, computed without retaining chunk content.
#[derive(Clone, Debug)]
pub struct SnapshotPlan {
    pub length: u64,
    /// SHA-256 of the whole file.
    pub checksum: String,
    pub offsets: Vec<u64>,
    pub sizes: Vec<usize>,
    pub objects: Vec<String>,
}

/// Reads the document sequentially with a bounded buffer; planning a large document needs only a
/// few hundred kilobytes of memory.
pub fn plan<R: Read + Seek>(stream: &mut R, format: ObjectFormat) -> Result<SnapshotPlan> {
    let length = stream.seek(SeekFrom::End(0))?;
    if length < 4 {
        return Err(Error::invalid("无法读取文档数据。"));
    }
    let mut overall = Sha256::new();
    let mut objects = Vec::new();
    let chunks = cut_chunks(stream, length, Some(&mut overall), Some(&mut |content: &[u8]| {
        objects.push(blob_id(format, content));
    }))?;
    let checksum = crate::hex(&overall.finalize());
    let sum: u64 = chunks.iter().map(|c| c.1 as u64).sum();
    if sum != length {
        return Err(Error::data("文档分块长度校验失败。"));
    }
    Ok(SnapshotPlan {
        length,
        checksum,
        offsets: chunks.iter().map(|c| c.0).collect(),
        sizes: chunks.iter().map(|c| c.1).collect(),
        objects,
    })
}

type ChunkSink<'a> = &'a mut dyn FnMut(&[u8]);

/// Reads the stream part by part in ascending order. `on_chunk` receives each finished chunk;
/// `overall` hashes every byte in file order.
fn cut_chunks<R: Read + Seek>(
    stream: &mut R,
    length: u64,
    mut overall: Option<&mut Sha256>,
    mut on_chunk: Option<ChunkSink<'_>>,
) -> Result<Vec<(u64, usize)>> {
    let boundaries = zip_boundaries(stream, length)?;
    let mut chunks = Vec::new();
    let mut buffer = vec![0u8; BUFFER];
    let mut chunk_buffer = if on_chunk.is_some() { vec![0u8; MAXIMUM_CHUNK] } else { Vec::new() };
    for part in 0..boundaries.len() - 1 {
        let end = boundaries[part + 1];
        let mut position = boundaries[part];
        stream.seek(SeekFrom::Start(position))?;
        let mut chunk_start = position;
        let mut chunk_length = 0usize;
        let mut rolling: u64 = 0;
        while position < end {
            let want = (BUFFER as u64).min(end - position) as usize;
            read_exact(stream, &mut buffer[..want])?;
            if let Some(hash) = overall.as_deref_mut() {
                hash.update(&buffer[..want]);
            }
            for &byte in &buffer[..want] {
                rolling = (rolling << 1).wrapping_add(GEAR[byte as usize]);
                if on_chunk.is_some() {
                    chunk_buffer[chunk_length] = byte;
                }
                chunk_length += 1;
                if chunk_length >= MINIMUM_CHUNK && ((rolling & 0xFFFF) == 0 || chunk_length >= MAXIMUM_CHUNK) {
                    chunks.push((chunk_start, chunk_length));
                    if let Some(sink) = on_chunk.as_deref_mut() {
                        sink(&chunk_buffer[..chunk_length]);
                    }
                    chunk_start = position + 1;
                    chunk_length = 0;
                    rolling = 0;
                }
                position += 1;
            }
        }
        if chunk_start < end {
            chunks.push((chunk_start, chunk_length));
            if let Some(sink) = on_chunk.as_deref_mut() {
                sink(&chunk_buffer[..chunk_length]);
            }
        }
    }
    Ok(chunks)
}

fn read_exact<R: Read>(stream: &mut R, buffer: &mut [u8]) -> Result<()> {
    stream.read_exact(buffer).map_err(|error| {
        if error.kind() == std::io::ErrorKind::UnexpectedEof {
            Error::data("文档在读取期间被修改。")
        } else {
            Error::Io(error)
        }
    })
}

fn u16_at(bytes: &[u8], at: usize) -> u64 {
    u16::from_le_bytes([bytes[at], bytes[at + 1]]) as u64
}

fn u32_at(bytes: &[u8], at: usize) -> u64 {
    u32::from_le_bytes([bytes[at], bytes[at + 1], bytes[at + 2], bytes[at + 3]]) as u64
}

/// Offsets that separate ZIP local headers, member payloads and the central directory. Any
/// structure outside the expectations (ZIP64, odd layouts) falls back to chunking the whole file.
fn zip_boundaries<R: Read + Seek>(stream: &mut R, length: u64) -> Result<Vec<u64>> {
    let fallback = vec![0, length];
    if length < 22 {
        return Ok(fallback);
    }
    let mut boundaries = BTreeSet::from([0, length]);
    let tail_length = (65_537u64 + 20).min(length) as usize;
    let mut tail = vec![0u8; tail_length];
    stream.seek(SeekFrom::Start(length - tail_length as u64))?;
    read_exact(stream, &mut tail)?;
    let mut eocd = None;
    let mut candidate = tail_length as i64 - 22;
    while candidate >= 0 {
        let at = candidate as usize;
        if u32_at(&tail, at) == 0x0605_4b50 && at + 22 + u16_at(&tail, at + 20) as usize == tail_length {
            eocd = Some(at);
            break;
        }
        candidate -= 1;
    }
    let Some(eocd) = eocd else { return Ok(fallback) };
    let end = length - tail_length as u64 + eocd as u64;
    let count = u16_at(&tail, eocd + 10);
    let directory = u32_at(&tail, eocd + 16);
    if directory > end || count == u16::MAX as u64 {
        return Ok(fallback);
    }
    let directory_length = end - directory;
    if directory_length > 4 * 1024 * 1024 {
        return Ok(fallback); // Pathological central directory; chunk the file as-is.
    }
    let mut records = vec![0u8; directory_length as usize];
    stream.seek(SeekFrom::Start(directory))?;
    read_exact(stream, &mut records)?;
    boundaries.insert(directory);
    let mut cursor = 0usize;
    let mut local_header = [0u8; 30];
    for _ in 0..count {
        if cursor + 46 > records.len() || u32_at(&records, cursor) != 0x0201_4b50 {
            return Ok(fallback);
        }
        let packed_size = u32_at(&records, cursor + 20);
        let local_offset = u32_at(&records, cursor + 42);
        if directory < 30 || local_offset > directory - 30 || packed_size > length {
            return Ok(fallback);
        }
        stream.seek(SeekFrom::Start(local_offset))?;
        read_exact(stream, &mut local_header)?;
        if u32_at(&local_header, 0) != 0x0403_4b50 {
            return Ok(fallback);
        }
        let payload = local_offset + 30 + u16_at(&local_header, 26) + u16_at(&local_header, 28);
        let payload_end = payload + packed_size;
        if payload > directory || payload_end > directory {
            return Ok(fallback);
        }
        boundaries.insert(local_offset);
        boundaries.insert(payload);
        boundaries.insert(payload_end);
        cursor += 46
            + u16_at(&records, cursor + 28) as usize
            + u16_at(&records, cursor + 30) as usize
            + u16_at(&records, cursor + 32) as usize;
    }
    Ok(boundaries.into_iter().collect())
}

/// Manifest text: format line, total length, whole-file SHA-256, then `<blob> <size>` per chunk.
pub fn manifest(length: u64, checksum: &str, sizes: &[usize], objects: &[String]) -> Result<Vec<u8>> {
    if sizes.len() != objects.len() {
        return Err(Error::data("分块写入数量不一致。"));
    }
    let mut text = format!("{FORMAT}\n{length}\n{checksum}\n");
    for (object, size) in objects.iter().zip(sizes) {
        text.push_str(&format!("{object} {size}\n"));
    }
    Ok(text.into_bytes())
}

/// A parsed manifest.
#[derive(Clone, Debug)]
pub struct DocumentSnapshot {
    pub length: u64,
    pub checksum: String,
    pub objects: Vec<String>,
    pub sizes: Vec<usize>,
}

fn digits(text: &str) -> Option<u64> {
    if text.is_empty() || !text.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    text.parse().ok()
}

impl DocumentSnapshot {
    pub fn read(bytes: &[u8]) -> Result<DocumentSnapshot> {
        let invalid = || Error::data("文档分块清单无效。");
        let text = std::str::from_utf8(bytes).map_err(|_| invalid())?;
        let lines: Vec<&str> = text.trim_end_matches('\n').split('\n').collect();
        if lines.len() < 4 || lines[0] != FORMAT {
            return Err(invalid());
        }
        let length = digits(lines[1]).filter(|&l| l >= 4).ok_or_else(invalid)?;
        let checksum = lines[2];
        if checksum.len() != 64 || !checksum.bytes().all(|b| b.is_ascii_digit() || (b'a'..=b'f').contains(&b)) {
            return Err(invalid());
        }
        let mut objects = Vec::with_capacity(lines.len() - 3);
        let mut sizes = Vec::with_capacity(lines.len() - 3);
        let mut sum = 0u64;
        for line in &lines[3..] {
            let damaged = || Error::data("文档分块清单损坏。");
            let mut parts = line.split(' ');
            let (Some(id), Some(size), None) = (parts.next(), parts.next(), parts.next()) else { return Err(damaged()) };
            let size = digits(size).ok_or_else(damaged)?;
            if !crate::is_object_id(id) || size == 0 || size > MAXIMUM_CHUNK as u64 {
                return Err(damaged());
            }
            objects.push(id.to_string());
            sizes.push(size as usize);
            sum += size;
        }
        if sum != length {
            return Err(Error::data("文档分块长度校验失败。"));
        }
        Ok(DocumentSnapshot { length, checksum: checksum.to_string(), objects, sizes })
    }

    /// Verifies and writes each chunk from `git cat-file --batch` output as it arrives, so memory
    /// stays bounded however large the document is.
    pub fn restore(&self, batch: &mut dyn Read, output: &mut dyn Write) -> Result<()> {
        let mut buffer = vec![0u8; BUFFER];
        let mut sha = Sha256::new();
        let mut reader = ByteReader::new(batch);
        for (object, &size) in self.objects.iter().zip(&self.sizes) {
            let header = reader.header(100).ok_or_else(|| Error::data("无法读取历史数据块。"))?;
            if header != format!("{object} blob {size}") {
                return Err(Error::data("历史数据块标识或长度校验失败。"));
            }
            let mut remaining = size;
            while remaining > 0 {
                let step = remaining.min(buffer.len());
                reader.read_exact(&mut buffer[..step]).map_err(|_| Error::data("历史数据块不完整。"))?;
                sha.update(&buffer[..step]);
                output.write_all(&buffer[..step])?;
                remaining -= step;
            }
            if reader.byte() != Some(b'\n') {
                return Err(Error::data("历史数据块不完整。"));
            }
        }
        if reader.byte().is_some() || crate::hex(&sha.finalize()) != self.checksum {
            return Err(Error::data("历史文档校验失败。"));
        }
        Ok(())
    }
}

/// Small buffered reader for `cat-file --batch` framing: header lines interleaved with raw bytes.
pub struct ByteReader<'a> {
    inner: &'a mut dyn Read,
    buffer: Vec<u8>,
    start: usize,
    end: usize,
}

impl<'a> ByteReader<'a> {
    pub fn new(inner: &'a mut dyn Read) -> Self {
        ByteReader { inner, buffer: vec![0u8; BUFFER], start: 0, end: 0 }
    }

    fn fill(&mut self) -> bool {
        if self.start < self.end {
            return true;
        }
        loop {
            match self.inner.read(&mut self.buffer) {
                Ok(0) => return false,
                Ok(count) => {
                    self.start = 0;
                    self.end = count;
                    return true;
                }
                Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
                Err(_) => return false,
            }
        }
    }

    pub fn byte(&mut self) -> Option<u8> {
        if !self.fill() {
            return None;
        }
        let value = self.buffer[self.start];
        self.start += 1;
        Some(value)
    }

    /// Reads one `\n`-terminated header line of at most `limit` bytes (lossy UTF-8).
    pub fn header(&mut self, limit: usize) -> Option<String> {
        let mut bytes = Vec::with_capacity(64);
        loop {
            let next = self.byte()?;
            if next == b'\n' {
                return Some(String::from_utf8_lossy(&bytes).into_owned());
            }
            if bytes.len() > limit {
                return None;
            }
            bytes.push(next);
        }
    }

    pub fn read_exact(&mut self, target: &mut [u8]) -> std::io::Result<()> {
        let mut filled = 0;
        while filled < target.len() {
            if !self.fill() {
                return Err(std::io::ErrorKind::UnexpectedEof.into());
            }
            let count = (self.end - self.start).min(target.len() - filled);
            target[filled..filled + count].copy_from_slice(&self.buffer[self.start..self.start + count]);
            self.start += count;
            filled += count;
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn gear_matches_reference_values() {
        // First entries of the C# CreateGear() table.
        let mut seed: u64 = 0x9E37_79B9_7F4A_7C15;
        seed ^= seed >> 12;
        seed ^= seed << 25;
        seed ^= seed >> 27;
        assert_eq!(GEAR[0], seed.wrapping_mul(2_685_821_657_736_338_717));
        assert_ne!(GEAR[1], GEAR[0]);
    }

    #[test]
    fn plain_bytes_cover_the_file_and_round_trip() {
        let mut data = Vec::new();
        let mut state: u32 = 7;
        for _ in 0..900_000 {
            state = state.wrapping_mul(1_103_515_245).wrapping_add(12_345);
            data.push((state >> 16) as u8);
        }
        let plan = plan(&mut Cursor::new(&data), ObjectFormat::Sha1).unwrap();
        assert!(plan.sizes.len() > 3);
        assert!(plan.sizes.iter().all(|&s| s <= MAXIMUM_CHUNK));
        assert_eq!(plan.sizes.iter().sum::<usize>(), data.len());
        let manifest = manifest(plan.length, &plan.checksum, &plan.sizes, &plan.objects).unwrap();
        let snapshot = DocumentSnapshot::read(&manifest).unwrap();
        // Emulate `git cat-file --batch` output and restore.
        let mut batch = Vec::new();
        for (index, (&offset, &size)) in plan.offsets.iter().zip(&plan.sizes).enumerate() {
            batch.extend_from_slice(format!("{} blob {}\n", plan.objects[index], size).as_bytes());
            batch.extend_from_slice(&data[offset as usize..offset as usize + size]);
            batch.push(b'\n');
        }
        let mut restored = Vec::new();
        snapshot.restore(&mut Cursor::new(batch), &mut restored).unwrap();
        assert_eq!(restored, data);
    }

    #[test]
    fn blob_ids_match_git() {
        // `printf hello | git hash-object --stdin`
        assert_eq!(blob_id(ObjectFormat::Sha1, b"hello"), "b6fc4c620b67d95f953a5c1c1230aaab5db5a1b0");
    }

    #[test]
    fn manifest_rejects_damage() {
        assert!(DocumentSnapshot::read(b"WORDGIT-SNAPSHOT 1\n4\nzz\nabc 4\n").is_err());
        let good = format!("WORDGIT-SNAPSHOT 1\n4\n{}\n{} 4\n", "a".repeat(64), "b".repeat(40));
        assert!(DocumentSnapshot::read(good.as_bytes()).is_ok());
        let wrong_sum = format!("WORDGIT-SNAPSHOT 1\n5\n{}\n{} 4\n", "a".repeat(64), "b".repeat(40));
        assert!(DocumentSnapshot::read(wrong_sum.as_bytes()).is_err());
    }
}
