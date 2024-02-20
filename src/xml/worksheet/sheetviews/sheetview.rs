use serde::{Deserialize, Serialize};
use crate::xml::worksheet::sheetviews::sheetview::pane::Pane;
use crate::xml::worksheet::sheetviews::sheetview::selection::Selection;

pub mod pane;
pub mod selection;

#[derive(Debug, Deserialize, Serialize, Default)]
pub(crate) struct SheetView {
    #[serde(rename = "@tabSelected", skip_serializing_if = "Option::is_none")]
    pub(crate) tab_selected: Option<u8>,
    #[serde(rename = "@zoomScale", skip_serializing_if = "Option::is_none")]
    pub(crate) zoom_scale: Option<u16>,
    #[serde(rename = "@topLeftCell", skip_serializing_if = "Option::is_none")]
    pub(crate) top_left_cell: Option<String>,
    #[serde(rename = "@workbookViewId")]
    workbook_view_id: u32,
    #[serde(rename = "pane", skip_serializing_if = "Option::is_none")]
    pane: Option<Vec<Pane>>,
    #[serde(rename = "selection", skip_serializing_if = "Option::is_none")]
    selection: Option<Vec<Selection>>,
}

impl SheetView {
    pub(crate) fn set_tab_selected(&mut self, tab_selected: u8) {
        self.tab_selected = Some(tab_selected);
    }

    pub(crate) fn set_zoom_scale(&mut self, zoom_scale: u16) {
        self.zoom_scale = Some(zoom_scale);
    }

    pub(crate) fn set_top_left_cell(&mut self, loc_ref: &str) {
        self.top_left_cell = Some(String::from(loc_ref));
    }

    pub(crate) fn set_selection(&mut self, loc_ref: String) {
        let selection = self.selection.take();
        let mut selection = selection.unwrap_or_else(|| vec![Selection::default()]);
        selection.last_mut().unwrap().add_selection(&loc_ref);
        self.selection = Some(selection);
    }

    pub(crate) fn set_freeze_panes(&mut self, from: &str, loc_ref: &str) {
        let pane = self.pane.take();
        let mut pane = pane.unwrap_or_else(|| vec![Pane::default()]);
        pane.last_mut().unwrap().set_freeze_panes(&loc_ref);
        self.top_left_cell = Some(from.to_string());
        self.pane = Some(pane);
    }
}