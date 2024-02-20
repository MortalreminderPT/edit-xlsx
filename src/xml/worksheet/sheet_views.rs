mod sheetview;

use serde::{Deserialize, Serialize};
use crate::api::cell::location::Location;
use self::sheetview::SheetView;

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct SheetViews {
    #[serde(rename = "sheetView")]
    pub(crate) sheet_view: Vec<SheetView>
}

impl Default for SheetViews {
    fn default() -> SheetViews {
        SheetViews {
            sheet_view: vec![SheetView::default()],
        }
    }
}

impl SheetViews {
    pub(crate) fn set_tab_selected(&mut self, tab_selected: u8) {
        self.sheet_view[0].set_tab_selected(tab_selected);
    }

    pub(crate) fn set_zoom_scale(&mut self, zoom_scale: u16) {
        self.sheet_view[0].set_zoom_scale(zoom_scale);
    }

    pub(crate) fn set_top_left_cell(&mut self, loc_ref: &str) {
        self.sheet_view[0].set_top_left_cell(loc_ref);
    }

    pub(crate) fn set_selection(&mut self, loc_ref: String) {
        self.sheet_view[0].set_selection(loc_ref);
    }

    pub(crate) fn set_freeze_panes<L: Location>(&mut self, loc: L) {
        self.sheet_view[0].set_freeze_panes(loc);
    }
}
