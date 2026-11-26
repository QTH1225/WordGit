//! DOCX container checks, plus the reader for the retired `wordgit.meta` storage format in which
//! each ZIP member was stored as its own blob. New commits use [`crate::snapshot`] chunks.

use crate::{Error, Result};
use base64::Engine;
use flate2::write::DeflateEncoder;
use flate2::Compression;
use sha2::{Digest, Sha256};
use std::fs::File;
use std::io::{Read, Seek, SeekFrom, Write};
use std::path::Path;

pub const META_NAME: &str = "wordgit.meta";
const MAGIC: &str = "wordgit 2";
const FINGERPRINT_PREFIX: &str = "fingerprint ";

/// Word already compresses these; the legacy format stored them uncompressed.
const STORED_EXTENSIONS: [&str; 19] = [
    ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf", ".bin", ".odttf", ".ttf",
    ".otf", ".woff", ".woff2", ".mp3", ".mp4", ".wav", ".zip",
];

/// Password-protected OOXML files are OLE compound files instead of ZIP archives.
pub fn is_encrypted(header: &[u8]) -> bool {
    header.len() >= 8 && header[..8] == [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]
}

/// Cheap structural check (signature and central directory only); used before committing and
/// before handing a restored file to the editor.
pub fn validate_file(path: &Path) -> Result<()> {
    let not_docx = || Error::data("文件不是可读取的 Word .docx。");
    let mut file = File::open(path).map_err(|error| {
        if error.kind() == std::io::ErrorKind::NotFound {
            Error::invalid(format!("找不到文件：{}", path.display()))
        } else {
            Error::Io(error)
        }
    })?;
    let mut header = [0u8; 8];
    let read = read_up_to(&mut file, &mut header)?;
    if read == header.len() && is_encrypted(&header) {
        return Err(Error::data("文档已加密或受密码保护，请先取消密码保护后再提交。"));
    }
    let names = zip_entry_names(&mut file).map_err(|_| not_docx())?;
    if !names.iter().any(|n| n == "[Content_Types].xml") || !names.iter().any(|n| n == "word/document.xml") {
        return Err(not_docx());
    }
    Ok(())
}

fn read_up_to(file: &mut File, buffer: &mut [u8]) -> std::io::Result<usize> {
    let mut read = 0;
    while read < buffer.len() {
        let count = file.read(&mut buffer[read..])?;
        if count == 0 {
            break;
        }
        read += count;
    }
    Ok(read)
}

/// Names of all members, read from the central directory (ZIP64 aware).
pub fn zip_entry_names<R: Read + Seek>(stream: &mut R) -> std::io::Result<Vec<String>> {
    let bad = || std::io::Error::new(std::io::ErrorKind::InvalidData, "not a zip archive");
    let length = stream.seek(SeekFrom::End(0))?;
    if length < 22 {
        return Err(bad());
    }
    let tail_length = (65_557u64).min(length) as usize;
    let mut tail = vec![0u8; tail_length];
    stream.seek(SeekFrom::Start(length - tail_length as u64))?;
    stream.read_exact(&mut tail)?;
    let eocd = (0..=tail_length - 22)
        .rev()
        .find(|&at| tail[at..at + 4] == [0x50, 0x4b, 0x05, 0x06])
        .ok_or_else(bad)?;
    let le16 = |b: &[u8], at: usize| u16::from_le_bytes([b[at], b[at + 1]]) as u64;
    let le32 = |b: &[u8], at: usize| u32::from_le_bytes(b[at..at + 4].try_into().unwrap()) as u64;
    let le64 = |b: &[u8], at: usize| u64::from_le_bytes(b[at..at + 8].try_into().unwrap());
    let mut count = le16(&tail, eocd + 10);
    let mut directory_size = le32(&tail, eocd + 12);
    let mut directory = le32(&tail, eocd + 16);
    if count == 0xFFFF || directory == 0xFFFF_FFFF || directory_size == 0xFFFF_FFFF {
        // ZIP64 end-of-central-directory locator sits right before the classic record.
        if eocd < 20 || tail[eocd - 20..eocd - 16] != [0x50, 0x4b, 0x06, 0x07] {
            return Err(bad());
        }
        let record = le64(&tail, eocd - 12);
        let mut zip64 = [0u8; 56];
        stream.seek(SeekFrom::Start(record))?;
        stream.read_exact(&mut zip64)?;
        if zip64[..4] != [0x50, 0x4b, 0x06, 0x06] {
            return Err(bad());
        }
        count = le64(&zip64, 32);
        directory_size = le64(&zip64, 40);
        directory = le64(&zip64, 48);
    }
    if directory_size > 64 * 1024 * 1024 || directory + directory_size > length {
        return Err(bad());
    }
    let mut records = vec![0u8; directory_size as usize];
    stream.seek(SeekFrom::Start(directory))?;
    stream.read_exact(&mut records)?;
    let mut names = Vec::new();
    let mut cursor = 0usize;
    for _ in 0..count {
        if cursor + 46 > records.len() || records[cursor..cursor + 4] != [0x50, 0x4b, 0x01, 0x02] {
            return Err(bad());
        }
        let name_length = le16(&records, cursor + 28) as usize;
        let extra = le16(&records, cursor + 30) as usize;
        let comment = le16(&records, cursor + 32) as usize;
        if cursor + 46 + name_length > records.len() {
            return Err(bad());
        }
        names.push(String::from_utf8_lossy(&records[cursor + 46..cursor + 46 + name_length]).into_owned());
        cursor += 46 + name_length + extra + comment;
    }
    Ok(names)
}

#[derive(Clone, Debug)]
pub struct MetaEntry {
    pub name: String,
    pub sha: String,
    pub size: u64,
}

#[derive(Clone, Debug)]
pub struct DocxMeta {
    pub fingerprint: String,
    pub entries: Vec<MetaEntry>,
}

/// Whole-document fingerprint over members ordered by name (ordinal).
pub fn fingerprint(entries: &[MetaEntry]) -> Result<String> {
    let mut ordered: Vec<&MetaEntry> = entries.iter().collect();
    ordered.sort_by(|a, b| ordinal_utf16(&a.name, &b.name));
    let mut outer = Sha256::new();
    outer.update(b"wordgit-docx/2\n");
    for entry in ordered {
        outer.update(entry.name.as_bytes());
        outer.update([0u8]);
        outer.update(entry.size.to_be_bytes());
        outer.update(from_hex(&entry.sha)?);
    }
    Ok(crate::hex(&outer.finalize()))
}

/// .NET's ordinal string comparison orders by UTF-16 code units.
fn ordinal_utf16(a: &str, b: &str) -> std::cmp::Ordering {
    a.encode_utf16().cmp(b.encode_utf16())
}

fn from_hex(text: &str) -> Result<Vec<u8>> {
    if !text.len().is_multiple_of(2) {
        return Err(Error::data("无法读取此仓库的版本元数据。"));
    }
    (0..text.len())
        .step_by(2)
        .map(|i| u8::from_str_radix(&text[i..i + 2], 16).map_err(|_| Error::data("无法读取此仓库的版本元数据。")))
        .collect()
}

fn is_hex64(value: &str) -> bool {
    value.len() == 64 && value.bytes().all(|b| b.is_ascii_digit() || (b'a'..=b'f').contains(&b))
}

pub fn parse_meta(text: &str) -> Result<DocxMeta> {
    let invalid = || Error::data("无法读取此仓库的版本元数据。");
    let lines: Vec<&str> = text.split('\n').collect();
    if lines.len() < 2 || lines[0] != MAGIC || !lines[1].starts_with(FINGERPRINT_PREFIX) {
        return Err(invalid());
    }
    let fingerprint = lines[1][FINGERPRINT_PREFIX.len()..].trim();
    if !is_hex64(fingerprint) {
        return Err(invalid());
    }
    let mut entries = Vec::new();
    for line in &lines[2..] {
        if line.is_empty() {
            continue;
        }
        let fields: Vec<&str> = line.split(' ').collect();
        if fields.len() != 3 {
            return Err(invalid());
        }
        let name_bytes = base64::engine::general_purpose::STANDARD.decode(fields[0]).map_err(|_| invalid())?;
        let name = String::from_utf8(name_bytes).map_err(|_| invalid())?;
        validate_name(&name)?;
        let size = if !fields[1].is_empty() && fields[1].bytes().all(|b| b.is_ascii_digit()) {
            fields[1].parse::<u64>().map_err(|_| invalid())?
        } else {
            return Err(invalid());
        };
        if !is_hex64(fields[2]) {
            return Err(invalid());
        }
        entries.push(MetaEntry { name, sha: fields[2].to_string(), size });
    }
    Ok(DocxMeta { fingerprint: fingerprint.to_string(), entries })
}

pub fn validate_name(name: &str) -> Result<()> {
    let bad = || Error::data("Word 文档中含有无法保存的部件名。");
    if name.is_empty() || name.encode_utf16().count() > 4096 || name.starts_with('/') || name.contains('\\') || name.contains('"') {
        return Err(bad());
    }
    let lower = name.to_lowercase();
    if lower == META_NAME || lower.starts_with(&format!("{META_NAME}/")) {
        return Err(Error::data("Word 文档包含与 WordGit 元数据同名的部件，暂时无法保存。"));
    }
    if name.chars().any(|c| (c as u32) < 0x20 || c as u32 == 0x7f) {
        return Err(bad());
    }
    if name.split('/').any(|segment| segment.is_empty() || segment == "." || segment == "..") {
        return Err(bad());
    }
    Ok(())
}

/// Whether the legacy writer stored this member without compression.
pub fn stored_uncompressed(name: &str) -> bool {
    // Same as .NET Path.GetExtension on the last path segment.
    let file = name.rsplit('/').next().unwrap_or(name);
    file.rfind('.')
        .map(|dot| file[dot..].to_lowercase())
        .is_some_and(|extension| STORED_EXTENSIONS.contains(&extension.as_str()))
}

/// Streaming ZIP writer (deflate or stored, data descriptors, fixed 1980-01-01 timestamps) used to
/// rebuild legacy expanded snapshots straight into the output.
pub struct ZipWriter<W: Write> {
    output: CountingWriter<W>,
    central: Vec<u8>,
    count: u64,
}

struct CountingWriter<W: Write> {
    inner: W,
    written: u64,
}

impl<W: Write> Write for CountingWriter<W> {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        let count = self.inner.write(buffer)?;
        self.written += count as u64;
        Ok(count)
    }
    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

/// One member being written; finish with [`ZipWriter::end_entry`].
pub struct ZipEntry {
    name: Vec<u8>,
    offset: u64,
    deflate: bool,
    crc: crc32fast::Hasher,
    size: u64,
    encoder: Option<DeflateEncoder<Vec<u8>>>,
    compressed: u64,
}

const DOS_DATE_1980: u16 = (1 << 5) | 1; // 1980-01-01
const DOS_TIME_MIDNIGHT: u16 = 0;

impl<W: Write> ZipWriter<W> {
    pub fn new(output: W) -> Self {
        ZipWriter { output: CountingWriter { inner: output, written: 0 }, central: Vec::new(), count: 0 }
    }

    pub fn start_entry(&mut self, name: &str) -> Result<ZipEntry> {
        let deflate = !stored_uncompressed(name);
        let name_bytes = name.as_bytes().to_vec();
        let offset = self.output.written;
        let mut header = Vec::with_capacity(30 + name_bytes.len());
        header.extend_from_slice(&0x0403_4b50u32.to_le_bytes());
        header.extend_from_slice(&20u16.to_le_bytes()); // version needed
        header.extend_from_slice(&(0x0008u16 | 0x0800).to_le_bytes()); // data descriptor, UTF-8 names
        header.extend_from_slice(&(if deflate { 8u16 } else { 0 }).to_le_bytes());
        header.extend_from_slice(&DOS_TIME_MIDNIGHT.to_le_bytes());
        header.extend_from_slice(&DOS_DATE_1980.to_le_bytes());
        header.extend_from_slice(&[0u8; 12]); // crc, sizes follow in the descriptor
        header.extend_from_slice(&(name_bytes.len() as u16).to_le_bytes());
        header.extend_from_slice(&0u16.to_le_bytes());
        header.extend_from_slice(&name_bytes);
        self.output.write_all(&header)?;
        Ok(ZipEntry {
            name: name_bytes,
            offset,
            deflate,
            crc: crc32fast::Hasher::new(),
            size: 0,
            encoder: deflate.then(|| DeflateEncoder::new(Vec::new(), Compression::default())),
            compressed: 0,
        })
    }

    pub fn write_entry(&mut self, entry: &mut ZipEntry, data: &[u8]) -> Result<()> {
        entry.crc.update(data);
        entry.size += data.len() as u64;
        match entry.encoder.as_mut() {
            Some(encoder) => {
                encoder.write_all(data)?;
                let pending = std::mem::take(encoder.get_mut());
                entry.compressed += pending.len() as u64;
                self.output.write_all(&pending)?;
            }
            None => {
                entry.compressed += data.len() as u64;
                self.output.write_all(data)?;
            }
        }
        Ok(())
    }

    pub fn end_entry(&mut self, mut entry: ZipEntry) -> Result<()> {
        if let Some(encoder) = entry.encoder.take() {
            let rest = encoder.finish()?;
            entry.compressed += rest.len() as u64;
            self.output.write_all(&rest)?;
        }
        if entry.size > u32::MAX as u64 || entry.compressed > u32::MAX as u64 || entry.offset > u32::MAX as u64 {
            return Err(Error::data("旧版文档部件过大，无法重建。"));
        }
        let crc = entry.crc.clone().finalize();
        let mut descriptor = Vec::with_capacity(16);
        descriptor.extend_from_slice(&0x0807_4b50u32.to_le_bytes());
        descriptor.extend_from_slice(&crc.to_le_bytes());
        descriptor.extend_from_slice(&(entry.compressed as u32).to_le_bytes());
        descriptor.extend_from_slice(&(entry.size as u32).to_le_bytes());
        self.output.write_all(&descriptor)?;
        let central = &mut self.central;
        central.extend_from_slice(&0x0201_4b50u32.to_le_bytes());
        central.extend_from_slice(&20u16.to_le_bytes()); // version made by
        central.extend_from_slice(&20u16.to_le_bytes());
        central.extend_from_slice(&(0x0008u16 | 0x0800).to_le_bytes());
        central.extend_from_slice(&(if entry.deflate { 8u16 } else { 0 }).to_le_bytes());
        central.extend_from_slice(&DOS_TIME_MIDNIGHT.to_le_bytes());
        central.extend_from_slice(&DOS_DATE_1980.to_le_bytes());
        central.extend_from_slice(&crc.to_le_bytes());
        central.extend_from_slice(&(entry.compressed as u32).to_le_bytes());
        central.extend_from_slice(&(entry.size as u32).to_le_bytes());
        central.extend_from_slice(&(entry.name.len() as u16).to_le_bytes());
        central.extend_from_slice(&[0u8; 8]); // extra, comment, disk, internal attributes
        central.extend_from_slice(&0u32.to_le_bytes()); // external attributes
        central.extend_from_slice(&(entry.offset as u32).to_le_bytes());
        central.extend_from_slice(&entry.name);
        self.count += 1;
        Ok(())
    }

    pub fn finish(mut self) -> Result<W> {
        let directory = self.output.written;
        if self.count > 0xFFFE || directory > u32::MAX as u64 {
            return Err(Error::data("旧版文档过大，无法重建。"));
        }
        let central = std::mem::take(&mut self.central);
        self.output.write_all(&central)?;
        let mut end = Vec::with_capacity(22);
        end.extend_from_slice(&0x0605_4b50u32.to_le_bytes());
        end.extend_from_slice(&[0u8; 4]);
        end.extend_from_slice(&(self.count as u16).to_le_bytes());
        end.extend_from_slice(&(self.count as u16).to_le_bytes());
        end.extend_from_slice(&(central.len() as u32).to_le_bytes());
        end.extend_from_slice(&(directory as u32).to_le_bytes());
        end.extend_from_slice(&0u16.to_le_bytes());
        self.output.write_all(&end)?;
        self.output.flush()?;
        Ok(self.output.inner)
    }
}

#[cfg(test)]
pub(crate) mod tests {
    use super::*;
    use std::io::Cursor;

    /// Builds a small DOCX-like archive for tests (members are deflated by our writer).
    pub fn sample_docx(text: &str, picture: &[u8]) -> Vec<u8> {
        let mut zip = ZipWriter::new(Vec::new());
        for (name, body) in [
            ("[Content_Types].xml", b"<Types/>".to_vec()),
            ("_rels/.rels", b"<Relationships/>".to_vec()),
            ("word/document.xml", format!("<w:document><w:t>{text}</w:t></w:document>").into_bytes()),
            ("word/media/image1.png", picture.to_vec()),
        ] {
            let mut entry = zip.start_entry(name).unwrap();
            zip.write_entry(&mut entry, &body).unwrap();
            zip.end_entry(entry).unwrap();
        }
        zip.finish().unwrap()
    }

    #[test]
    fn writer_output_is_a_readable_zip() {
        let bytes = sample_docx("你好", &[1, 2, 3]);
        let names = zip_entry_names(&mut Cursor::new(&bytes)).unwrap();
        assert_eq!(names, ["[Content_Types].xml", "_rels/.rels", "word/document.xml", "word/media/image1.png"]);
        let path = std::env::temp_dir().join(format!("wg-docx-{}.docx", crate::random_hex(4)));
        std::fs::write(&path, &bytes).unwrap();
        validate_file(&path).unwrap();
        std::fs::write(&path, b"not a zip at all, just text bytes").unwrap();
        assert!(validate_file(&path).is_err());
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn meta_round_trip_and_names() {
        let entry = MetaEntry { name: "word/document.xml".into(), sha: "a".repeat(64), size: 5 };
        let print = fingerprint(std::slice::from_ref(&entry)).unwrap();
        let text = format!(
            "wordgit 2\nfingerprint {print}\n{} 5 {}\n",
            base64::engine::general_purpose::STANDARD.encode("word/document.xml"),
            "a".repeat(64)
        );
        let meta = parse_meta(&text).unwrap();
        assert_eq!(meta.entries.len(), 1);
        assert_eq!(meta.fingerprint, print);
        assert!(validate_name("../evil").is_err());
        assert!(validate_name("wordgit.meta").is_err());
        assert!(stored_uncompressed("word/media/a.PNG"));
        assert!(!stored_uncompressed("word/document.xml"));
    }
}
