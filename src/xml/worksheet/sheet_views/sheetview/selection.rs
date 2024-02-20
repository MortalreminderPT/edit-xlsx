use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Selection {
    #[serde(rename = "@pane", skip_serializing_if = "Option::is_none")]
    pane: Option<String>,
    #[serde(rename = "@activeCell", skip_serializing_if = "Option::is_none")]
    active_cell: Option<String>,
    #[serde(rename = "@sqref", skip_serializing_if = "Option::is_none")]
    sqref: Option<String>,
}

impl Default for Selection {
    fn default() -> Self {
        Self {
            active_cell: Some(String::new()),
            sqref: Some(String::new()),
            pane: None,
        }
    }
}

impl Selection {
    pub(crate) fn add_selection(&mut self, loc_ref: &str) {
        let mut sqref = self.sqref.take().unwrap_or_default();
        sqref.push_str(&format!(" {}", loc_ref));
        self.sqref = Some(sqref);
    }
}
