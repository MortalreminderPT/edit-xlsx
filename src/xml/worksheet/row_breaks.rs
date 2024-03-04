use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct RowBreaks {
    #[serde(rename = "@count", default)]
    count: u32,
    #[serde(rename = "@manualBreakCount", default)]
    manual_break_count: u32,
    #[serde(rename = "brk", default, skip_serializing_if = "Vec::is_empty")]
    brk: Vec<Break>,
}

#[derive(Debug, Deserialize, Serialize)]
struct Break {
    #[serde(rename = "@id", skip_serializing_if = "Option::is_none")]
    id: Option<u32>,
    #[serde(rename = "@max", skip_serializing_if = "Option::is_none")]
    max: Option<u32>,
    #[serde(rename = "@man", skip_serializing_if = "Option::is_none")]
    man: Option<u32>,
}