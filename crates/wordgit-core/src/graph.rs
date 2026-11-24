//! Lane layout of the commit graph (one row per revision, in `--topo-order`). The add-in draws
//! the rows; this keeps the layout identical to the original Windows add-in.

use crate::history::Revision;
use serde::Serialize;
use std::collections::{HashMap, HashSet};

#[derive(Clone, Debug, Serialize)]
pub struct GraphEdge {
    pub from: usize,
    pub to: usize,
    pub color: usize,
    /// Vertical anchors in half-row units: 0 top, 1 middle (the node), 2 bottom.
    pub start: u8,
    pub end: u8,
}

#[derive(Clone, Debug, Serialize)]
pub struct GraphRow {
    pub node: usize,
    pub node_color: usize,
    pub columns: usize,
    pub is_head: bool,
    pub edges: Vec<GraphEdge>,
}

#[derive(Clone)]
struct Lane {
    id: String,
    color: usize,
    serial: usize,
}

pub fn layout(revisions: &[Revision], head: Option<&str>) -> Vec<GraphRow> {
    let mut result = Vec::with_capacity(revisions.len());
    let mut remaining: HashSet<&str> = revisions.iter().map(|r| r.id.as_str()).collect();
    let mut lanes: Vec<Lane> = Vec::new();
    let mut next_color = 0;
    let mut serial = 0;
    for revision in revisions {
        remaining.remove(revision.id.as_str());
        let incoming = lanes.clone();
        let node_index = match lanes.iter().position(|lane| lane.id == revision.id) {
            Some(index) => index,
            None => {
                lanes.push(Lane { id: revision.id.clone(), color: next_color, serial });
                next_color += 1;
                serial += 1;
                lanes.len() - 1
            }
        };
        let node = lanes[node_index].clone();
        let mut row = GraphRow {
            node: node_index,
            node_color: node.color,
            columns: 0,
            is_head: head == Some(revision.id.as_str()),
            edges: Vec::new(),
        };
        let mut parents: Vec<&str> = Vec::new();
        for parent in &revision.parents {
            if remaining.contains(parent.as_str()) && !parents.contains(&parent.as_str()) {
                parents.push(parent);
            }
        }
        lanes.remove(node_index);
        for (index, parent) in parents.iter().enumerate() {
            if lanes.iter().any(|lane| lane.id == *parent) {
                continue;
            }
            let insert_at = (node_index + index).min(lanes.len());
            let color = if index == 0 {
                node.color
            } else {
                next_color += 1;
                next_color - 1
            };
            lanes.insert(insert_at, Lane { id: parent.to_string(), color, serial });
            serial += 1;
        }
        let positions: HashMap<usize, usize> = lanes.iter().enumerate().map(|(i, lane)| (lane.serial, i)).collect();
        for (index, lane) in incoming.iter().enumerate() {
            if lane.serial == node.serial {
                row.edges.push(GraphEdge { from: index, to: node_index, color: lane.color, start: 0, end: 1 });
            } else {
                row.edges.push(GraphEdge { from: index, to: positions[&lane.serial], color: lane.color, start: 0, end: 2 });
            }
        }
        for parent in &parents {
            let index = lanes.iter().position(|lane| lane.id == *parent).expect("parent lane exists");
            row.edges.push(GraphEdge { from: node_index, to: index, color: lanes[index].color, start: 1, end: 2 });
        }
        row.columns = (node_index + 1).max(incoming.len()).max(lanes.len());
        result.push(row);
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    fn revision(id: &str, parents: &[&str]) -> Revision {
        Revision {
            id: id.into(),
            parents: parents.iter().map(|p| p.to_string()).collect(),
            date: String::new(),
            author: String::new(),
            references: Vec::new(),
            subject: String::new(),
            message: String::new(),
        }
    }

    #[test]
    fn merge_opens_and_closes_a_lane() {
        // m merges b into a; both descend from r.
        let revisions = vec![revision("m", &["a", "b"]), revision("b", &["r"]), revision("a", &["r"]), revision("r", &[])];
        let rows = layout(&revisions, Some("m"));
        assert!(rows[0].is_head);
        assert_eq!(rows[0].columns, 2);
        assert_eq!(rows[1].node, 1);
        assert_eq!(rows[3].node, 0);
        assert_eq!(rows[3].columns, 1);
        // a's parent r already has a lane (opened by b): a joins it instead of opening a new one.
        let joins: Vec<_> = rows[2].edges.iter().filter(|e| e.start == 1).collect();
        assert_eq!(joins.len(), 1);
        assert_eq!(joins[0].to, 0);
    }
}
