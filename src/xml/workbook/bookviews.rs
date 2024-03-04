use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct BookViews {
    #[serde(rename = "workbookView")]
    pub(crate) book_views: Vec<WorkbookView>
}

#[derive(Debug, Deserialize, Serialize, Default)]
pub(crate) struct WorkbookView {
    #[serde(rename = "@xWindow", default)]
    x_window: i32,
    #[serde(rename = "@yWindow", default)]
    y_window: i32,
    #[serde(rename = "@windowWidth", default)]
    pub(crate) window_width: u32,
    #[serde(rename = "@windowHeight", default)]
    pub(crate) window_height: u32,
    #[serde(rename = "@tabRatio", skip_serializing_if = "Option::is_none")]
    pub(crate) tab_ratio: Option<u32>,
    #[serde(rename = "@activeTab", skip_serializing_if = "Option::is_none")]
    pub(crate) active_tab: Option<u32>
}

impl Default for BookViews {
    fn default() -> Self {
        BookViews {
            book_views: vec![Default::default()],
        }
    }
}

impl BookViews {
    pub(crate) fn set_tab_ratio(&mut self, tab_ratio: u32) {
        self.book_views[0].tab_ratio = Some(tab_ratio);
    }

    pub(crate) fn set_active_tab(&mut self, active_tab: u32) {
        self.book_views[0].active_tab = Some(active_tab);
    }
}