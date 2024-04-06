mod sheetview;

use serde::{Deserialize, Serialize};
use crate::api::cell::location::{Location, LocationRange};
use self::sheetview::SheetView;

#[derive(Debug, Clone, Deserialize, Serialize)]
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
    pub(crate) fn set_right_to_left(&mut self, right_to_left: u8) {
        self.sheet_view[0].set_right_to_left(right_to_left);
    }
    
    pub(crate) fn set_tab_selected(&mut self, tab_selected: u8) {
        self.sheet_view[0].set_tab_selected(tab_selected);
    }

    pub(crate) fn set_zoom_scale(&mut self, zoom_scale: u16) {
        self.sheet_view[0].set_zoom_scale(zoom_scale);
    }

    pub(crate) fn set_top_left_cell(&mut self, loc_ref: &str) {
        self.sheet_view[0].set_top_left_cell(loc_ref);
    }

    pub(crate) fn set_selection<L: LocationRange>(&mut self, loc_ref: &L) {
        self.sheet_view[0].set_selection(loc_ref);
    }

    pub(crate) fn set_freeze_panes<L: Location>(&mut self, loc: L) {
        self.sheet_view[0].set_frozen_panes(loc, true);
    }

    pub(crate) fn split_panes(&mut self, x_split: u32, y_split: u32) {
        self.sheet_view[0].set_panes(x_split, y_split);
    }
}
