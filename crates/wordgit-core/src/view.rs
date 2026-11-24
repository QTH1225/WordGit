//! Everything the history pane shows, read with as few Git processes as possible.

use crate::history::{Repo, Revision, HEADS};
use crate::remote::{RemoteState, REMOTE_PREFIX};
use crate::{object_id, Error, Result};
use serde::Serialize;
use std::collections::BTreeMap;

#[derive(Clone, Debug, Serialize)]
pub struct RepositoryView {
    pub remote: RemoteState,
    /// Local branches, then remote-tracking references when a remote is linked.
    pub branches: Vec<String>,
    pub history: Vec<Revision>,
}

impl Repo {
    /// All branch tips and the relevant configuration in two Git processes, plus the history.
    pub fn read_view(&self) -> Result<RepositoryView> {
        if !self.exists() {
            let remote = RemoteState {
                local_branch: "main".into(),
                remote_branch: "main".into(),
                remote_reference: format!("{REMOTE_PREFIX}main"),
                ..Default::default()
            }
            .finish();
            return Ok(RepositoryView { remote, branches: Vec::new(), history: Vec::new() });
        }
        let mut tips = BTreeMap::new();
        let mut local = None;
        let references =
            self.text(&["for-each-ref", "--sort=refname", "--format=%(HEAD)%00%(refname)%00%(objectname)", HEADS, REMOTE_PREFIX])?;
        for line in references.split(['\r', '\n']).filter(|l| !l.is_empty()) {
            let fields: Vec<&str> = line.split('\0').collect();
            if fields.len() != 3 {
                return Err(Error::data("无法读取仓库分支。"));
            }
            tips.insert(fields[1].to_string(), object_id(fields[2])?);
            if fields[0].trim() == "*" {
                if let Some(name) = fields[1].strip_prefix(HEADS) {
                    local = Some(name.to_string());
                }
            }
        }
        let local = match local {
            Some(local) => local,
            None => self.active_branch()?, // An unborn branch has no ref yet.
        };
        let mut settings = BTreeMap::new();
        let config = self.invoke(
            None,
            None,
            &["config", "--local", "--null", "--get-regexp", "^(remote\\.origin\\.url|branch\\..*\\.merge)$"],
        )?;
        if config.code != 1 {
            let bytes = config.checked()?.stdout;
            for entry in String::from_utf8_lossy(&bytes).split('\0').filter(|e| !e.is_empty()) {
                if let Some((key, value)) = entry.split_once('\n') {
                    settings.insert(key.to_string(), value.to_string());
                }
            }
        }
        let merge = settings.get(&format!("branch.{local}.merge"));
        let remote_branch = merge.and_then(|m| m.strip_prefix(HEADS)).unwrap_or(&local).to_string();
        let mut state = RemoteState {
            url: settings.get("remote.origin.url").cloned(),
            local_tip: tips.get(&format!("{HEADS}{local}")).cloned(),
            remote_reference: format!("{REMOTE_PREFIX}{remote_branch}"),
            local_branch: local,
            remote_branch,
            ..Default::default()
        };
        if state.url.is_some() {
            state.remote_tip = tips.get(&state.remote_reference).cloned();
        }
        let state = self.complete_sync_state(state)?;
        let mut branches: Vec<String> = tips.keys().filter_map(|r| r.strip_prefix(HEADS)).map(str::to_string).collect();
        if state.linked {
            let head = format!("{REMOTE_PREFIX}HEAD");
            branches.extend(tips.keys().filter(|r| r.starts_with(REMOTE_PREFIX) && **r != head).cloned());
        }
        Ok(RepositoryView { remote: state, branches, history: self.history(None)? })
    }
}
