//! WordGit core.
//!
//! A Word/WPS document's history lives in a bare Git repository outside the document's folder.
//! Each commit stores the DOCX as content-defined chunks (see [`snapshot`]), so successive
//! versions share unchanged pictures and XML parts. Every Git invocation goes through [`git`],
//! which isolates the user's configuration and stops whole process trees on timeout.
//!
//! The on-disk formats are shared with the original Windows C# add-in: repositories written by
//! either implementation are readable by the other, and committing the same file produces the
//! same snapshot tree.

pub mod docx;
pub mod fsutil;
pub mod git;
pub mod graph;
pub mod history;
pub mod journal;
pub mod maintenance;
pub mod paths;
pub mod remote;
pub mod review;
pub mod snapshot;
pub mod view;

pub use history::{PreparedSnapshot, Repo, Revision};

/// Errors carry user-facing Chinese messages, matching the original add-in's wording.
#[derive(Debug, thiserror::Error)]
pub enum Error {
    /// The request cannot be carried out in the current state (stale tips, bad input, ...).
    #[error("{0}")]
    Invalid(String),
    /// Stored or supplied data is corrupt or inconsistent.
    #[error("{0}")]
    Data(String),
    #[error("{0}")]
    Timeout(String),
    #[error("操作已取消。")]
    Cancelled,
    #[error("{0}")]
    Io(#[from] std::io::Error),
}

impl Error {
    pub fn invalid(message: impl Into<String>) -> Self {
        Error::Invalid(message.into())
    }

    pub fn data(message: impl Into<String>) -> Self {
        Error::Data(message.into())
    }

    /// Stable identifier for API clients.
    pub fn kind(&self) -> &'static str {
        match self {
            Error::Invalid(_) => "invalid",
            Error::Data(_) => "data",
            Error::Timeout(_) => "timeout",
            Error::Cancelled => "cancelled",
            Error::Io(_) => "io",
        }
    }
}

pub type Result<T> = std::result::Result<T, Error>;

/// Lowercase hexadecimal encoding.
pub fn hex(bytes: &[u8]) -> String {
    const DIGITS: &[u8; 16] = b"0123456789abcdef";
    let mut out = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        out.push(DIGITS[(byte >> 4) as usize] as char);
        out.push(DIGITS[(byte & 15) as usize] as char);
    }
    out
}

/// Accepts a full SHA-1 or SHA-256 Git object id.
pub fn is_object_id(candidate: &str) -> bool {
    (candidate.len() == 40 || candidate.len() == 64)
        && candidate.bytes().all(|b| b.is_ascii_digit() || (b'a'..=b'f').contains(&b))
}

/// Validates a Git object id, returning it unchanged.
pub fn object_id(candidate: &str) -> Result<String> {
    if is_object_id(candidate) {
        Ok(candidate.to_string())
    } else {
        Err(Error::invalid("无效的 Git 版本标识。"))
    }
}

/// Random lowercase hex string of `bytes` random bytes.
pub fn random_hex(bytes: usize) -> String {
    let mut buffer = vec![0u8; bytes];
    getrandom::fill(&mut buffer).expect("system random source unavailable");
    hex(&buffer)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn object_ids() {
        assert!(is_object_id(&"a".repeat(40)));
        assert!(is_object_id(&"0".repeat(64)));
        assert!(!is_object_id(&"A".repeat(40)));
        assert!(!is_object_id(&"a".repeat(41)));
        assert_eq!(hex(&[0, 15, 255]), "000fff");
    }
}
