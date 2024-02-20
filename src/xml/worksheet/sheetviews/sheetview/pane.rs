use serde::{Deserialize, Serialize};

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
    pub(crate) fn set_freeze_panes(&mut self, loc_ref: &str) {
        let mut top_left_cell = self.top_left_cell.take().unwrap_or_default();
        top_left_cell = loc_ref.to_string();
        self.x_split = Some(1);
        self.y_split = Some(1);
        self.state = Some("frozen".to_string());
        self.top_left_cell = Some(top_left_cell);
    }
}