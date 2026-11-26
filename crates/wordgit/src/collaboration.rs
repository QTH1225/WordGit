use serde::Deserialize;
use serde_json::{json, Value};
use wordgit_core::{Error, Result};

#[derive(Deserialize)]
pub struct CollaborationRequest {
    local: Vec<Edit>,
    remote: Vec<Edit>,
    comparison: Vec<Edit>,
}

#[derive(Deserialize)]
struct Edit {
    story: String,
    kind: u32,
    start: usize,
    end: usize,
    text: String,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Span(usize, usize);

impl Span {
    fn intersects(self, other: Self) -> bool {
        if self.0 == self.1 {
            return self.0 == other.0 || other.0 < self.0 && self.0 < other.1;
        }
        other.0 < self.1 && self.0 < other.1 || other.0 == other.1 && other.0 == self.0
    }

    fn union(self, other: Self) -> Self {
        Self(self.0.min(other.0), self.1.max(other.1))
    }
}

struct Change<'a> {
    story: &'a str,
    base: Span,
    revised: Span,
    end: usize,
    marks: Vec<usize>,
    text_edit: bool,
    before: String,
    after: String,
}

fn inserted(kind: u32) -> bool {
    matches!(kind, 1 | 15)
}
fn deleted(kind: u32) -> bool {
    matches!(kind, 2 | 14)
}

// Office keeps deleted text and inserted text in the same range coordinate system.
// Each group is projected to both clean copies, and a replacement is classified together.
fn changes(edits: &[Edit]) -> Vec<Change<'_>> {
    let mut ordered: Vec<_> = edits.iter().enumerate().collect();
    ordered.sort_by_key(|&(_, edit)| (edit.story.as_str(), edit.start, edit.end));
    let mut result: Vec<Change<'_>> = Vec::new();
    for story in ordered.chunk_by(|a, b| a.1.story == b.1.story) {
        let omit = |kind: fn(u32) -> bool| {
            let mut ranges: Vec<Span> = Vec::new();
            for &(_, edit) in story.iter().filter(|(_, e)| kind(e.kind)) {
                let range = Span(edit.start, edit.end);
                if let Some(last) = ranges.last_mut().filter(|last| range.0 <= last.1) {
                    *last = last.union(range);
                } else {
                    ranges.push(range);
                }
            }
            ranges
        };
        let base_omissions = omit(inserted);
        let revised_omissions = omit(deleted);
        let project = |position: usize, ranges: &[Span]| {
            position.saturating_sub(
                ranges
                    .iter()
                    .take_while(|range| range.0 < position)
                    .map(|range| position.min(range.1) - range.0)
                    .sum(),
            )
        };
        for &(index, edit) in story {
            let base = Span(
                project(edit.start, &base_omissions),
                project(edit.end, &base_omissions),
            );
            let revised = Span(
                project(edit.start, &revised_omissions),
                project(edit.end, &revised_omissions),
            );
            let text_edit = inserted(edit.kind) || deleted(edit.kind);
            let group = result.last_mut().filter(|last| {
                last.story == edit.story
                    && last.text_edit
                    && text_edit
                    && last.end == edit.start
                    && last.base.1 >= base.0
                    && last.revised.1 >= revised.0
            });
            let change = if let Some(last) = group {
                last.base = last.base.union(base);
                last.revised = last.revised.union(revised);
                last.end = edit.end;
                last
            } else {
                result.push(Change {
                    story: &edit.story,
                    base,
                    revised,
                    end: edit.end,
                    text_edit,
                    marks: Vec::new(),
                    before: String::new(),
                    after: String::new(),
                });
                result.last_mut().unwrap()
            };
            change.marks.push(index);
            if deleted(edit.kind) {
                change.before.push_str(&edit.text);
            } else {
                change.after.push_str(&edit.text);
            }
        }
    }
    result
}

struct Conflict {
    local: Vec<usize>,
    remote: Vec<usize>,
}

pub fn plan(request: &CollaborationRequest) -> Result<Value> {
    if [&request.local, &request.remote, &request.comparison]
        .into_iter()
        .flatten()
        .any(|e| e.start > e.end)
    {
        return Err(Error::invalid("编辑器返回了无效的修订范围。"));
    }
    let local = changes(&request.local);
    let remote = changes(&request.remote);
    let mut conflicts: Vec<Conflict> = Vec::new();
    for (i, ours) in local.iter().enumerate() {
        for (j, theirs) in remote.iter().enumerate() {
            if ours.story != theirs.story || !ours.base.intersects(theirs.base) {
                continue;
            }
            if ours.base == theirs.base
                && ours.before == theirs.before
                && ours.after == theirs.after
            {
                continue;
            }
            let mut conflict = Conflict {
                local: vec![i],
                remote: vec![j],
            };
            let mut n = 0;
            while n < conflicts.len() {
                if conflicts[n].local.contains(&i) || conflicts[n].remote.contains(&j) {
                    let joined = conflicts.remove(n);
                    conflict.local.extend(joined.local);
                    conflict.remote.extend(joined.remote);
                } else {
                    n += 1;
                }
            }
            conflict.local.sort_unstable();
            conflict.local.dedup();
            conflict.remote.sort_unstable();
            conflict.remote.dedup();
            conflicts.push(conflict);
        }
    }
    let mut decisions = vec![Value::Null; request.comparison.len()];
    let matching = |edit: &Change<'_>, candidate: &Change<'_>, local_side| {
        let side = if local_side { edit.base } else { edit.revised };
        edit.story == candidate.story
            && edit.text_edit == candidate.text_edit
            && side.intersects(candidate.revised)
    };
    for edit in changes(&request.comparison) {
        let left: Vec<_> = local
            .iter()
            .enumerate()
            .filter(|(_, c)| matching(&edit, c, true))
            .map(|(i, _)| i)
            .collect();
        let right: Vec<_> = remote
            .iter()
            .enumerate()
            .filter(|(_, c)| matching(&edit, c, false))
            .map(|(i, _)| i)
            .collect();
        let related: Vec<_> = conflicts
            .iter()
            .enumerate()
            .filter(|(_, c)| {
                c.local.iter().any(|i| left.contains(i))
                    || c.remote.iter().any(|i| right.contains(i))
            })
            .collect();
        let decision = match related.as_slice() {
            [] if left.is_empty() => json!({ "incoming": !right.is_empty() }),
            [] if right.is_empty() => json!({ "incoming": false }),
            [(id, conflict)]
                if left.iter().all(|i| conflict.local.contains(i))
                    && right.iter().all(|i| conflict.remote.contains(i)) =>
            {
                json!({ "conflict": id })
            }
            _ => {
                return Err(Error::invalid(
                    "编辑器把几处改动合成了一条无法安全拆开的修订，主稿与远端历史均已保留。",
                ))
            }
        };
        for i in edit.marks {
            decisions[i] = decision.clone();
        }
    }
    let description = |indices: &[usize], side: &[Change<'_>]| {
        indices
            .iter()
            .map(|&i| {
                let edit = &side[i];
                if edit.after.is_empty() && !edit.before.is_empty() {
                    format!("删除「{}」", edit.before)
                } else if edit.after.is_empty() {
                    String::from("格式或结构变更")
                } else {
                    edit.after.clone()
                }
            })
            .collect::<Vec<_>>()
            .join("\n")
    };
    let conflicts: Vec<_> = conflicts.iter().map(|conflict| json!({
        "local": description(&conflict.local, &local),
        "remote": description(&conflict.remote, &remote),
        "original": conflict.local.iter().filter_map(|&i| (!local[i].before.is_empty()).then_some(local[i].before.as_str()))
            .collect::<Vec<_>>().join("\n"),
    })).collect();
    Ok(json!({ "conflicts": conflicts, "decisions": decisions }))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn groups_replacements_and_keeps_unrelated_edits() {
        let make = |kind, start, end, text: &str| Edit {
            story: "1:0".into(),
            kind,
            start,
            end,
            text: text.into(),
        };
        let result = plan(&CollaborationRequest {
            local: vec![
                make(1, 5, 6, "A"),
                make(2, 21, 22, "0"),
                make(1, 22, 23, "2"),
            ],
            remote: vec![
                make(1, 10, 11, "B"),
                make(2, 21, 22, "0"),
                make(1, 22, 23, "5"),
            ],
            comparison: vec![
                make(2, 5, 6, "A"),
                make(1, 11, 12, "B"),
                make(2, 22, 23, "2"),
                make(1, 23, 24, "5"),
            ],
        })
        .unwrap();
        assert_eq!(
            result["conflicts"],
            json!([{"local":"2","remote":"5","original":"0"}])
        );
        assert_eq!(
            result["decisions"],
            json!([
                {"incoming":false},{"incoming":true},{"conflict":0},{"conflict":0}
            ])
        );
    }
}
