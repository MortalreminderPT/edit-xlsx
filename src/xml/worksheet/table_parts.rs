use serde::{Deserialize, Serialize};
use crate::api::relationship::Rel;

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
pub(crate) struct TableParts {
    #[serde(rename = "@count", default)]
    count: u32,
    #[serde(rename = "tablePart", default)]
    table_part: Vec<TablePart>,
}

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
pub(crate) struct TablePart {
    #[serde(rename(serialize = "@r:id", deserialize = "@id"))]
    r_id: Rel,
}