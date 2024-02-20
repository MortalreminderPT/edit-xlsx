use serde::{Deserialize, Serialize};
use crate::api::cell::location::Location;
use crate::xml::worksheet::sheet_views::sheetview::selection::ActivePane;

#[derive(Debug, Deserialize, Serialize, Default)]
pub(crate) struct Pane {
    #[serde(rename = "@xSplit", skip_serializing_if = "Option::is_none")]
    x_split: Option<u32>,
    #[serde(rename = "@ySplit", skip_serializing_if = "Option::is_none")]
    y_split: Option<u32>,
    #[serde(rename = "@topLeftCell", skip_serializing_if = "Option::is_none")]
    top_left_cell: Option<String>,
    #[serde(rename = "@activePane", skip_serializing_if = "Option::is_none")]
    active_pane: Option<String>,
    #[serde(rename = "@state", skip_serializing_if = "Option::is_none")]
    state: Option<String>,
}

impl Pane {
    pub(crate) fn from_location<L: Location>(loc: L) -> Pane {
        let (row, col) = loc.to_location();
        Pane {
            x_split: Some(row - 1),
            y_split: Some(col - 1),
            top_left_cell: Some(loc.to_ref()),
            active_pane: Some(String::from("bottomRight")),
            state: Some(String::from("frozen")),
        }
    }

    pub(crate) fn default_pane<L: Location>(active_pane: ActivePane<L>) -> Self {
        let (y_split, x_split) = active_pane.get_split();
        Self {
            x_split,
            y_split,
            top_left_cell: active_pane.get_sqref(),
            active_pane: Some(String::from(active_pane.get_pane())),
            state: Some(String::from("frozen")),
        }
    }
}