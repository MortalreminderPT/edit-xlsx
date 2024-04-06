use serde::{Deserialize, Serialize};
use crate::api::cell::location::Location;
use crate::xml::worksheet::sheet_views::sheetview::selection::ActivePane;

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
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
    pub(crate) fn from_location<L: Location>(loc: L, active_pane: &str) -> Pane {
        let (row, col) = loc.to_location();
        let (y_split, x_split) = (if row == 1 { None } else { Some(row - 1) }, if col == 1 { None } else { Some(col - 1) });
        Pane {
            x_split,
            y_split,
            top_left_cell: Some(loc.to_ref()),
            active_pane: Some(String::from(active_pane)),
            state: Some(String::from("frozen")),
        }
    }
    
    pub(crate) fn from_split(x_split: u32, y_split: u32) -> Self {
        Self {
            x_split: Some(x_split),
            y_split: Some(y_split),
            top_left_cell: None,
            active_pane: Some("bottomRight".to_string()),
            state: None,
        }
    }
}