use serde::{Deserialize, Serialize};
use crate::utils::col_helper;

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
pub(crate) struct MergeCells {
    #[serde(rename = "@count", default)]
    count: u32,
    #[serde(rename = "mergeCell", default)]
    merge_cell: Vec<MergeCell>
}

impl MergeCells {
    pub(crate) fn add_merge_cell(&mut self, first_row: u32, first_col: u32, last_row: u32, last_col: u32) {
        let merge_cell = MergeCell::from_range(first_row, first_col, last_row, last_col);
        self.merge_cell.push(merge_cell);
        self.count += 1;
    }
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct MergeCell {
    #[serde(rename = "@ref")]
    cell_ref: String
}

impl MergeCell {
    pub(crate) fn from_range(first_row: u32, first_col: u32, last_row: u32, last_col: u32) -> MergeCell {
        let first_col = col_helper::to_col_name(first_col);
        let last_col = col_helper::to_col_name(last_col);
        MergeCell {
            cell_ref: format!("{first_col}{first_row}:{last_col}{last_row}")
        }
    }
}