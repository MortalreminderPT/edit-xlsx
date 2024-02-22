use serde::{Deserialize, Serialize};
use crate::api::cell::location::{Location, LocationRange};
use crate::xml::worksheet::sheet_data::cell::Sqref;
use crate::xml::worksheet::sheet_views::sheetview::pane::Pane;
use crate::xml::worksheet::sheet_views::sheetview::selection::{Selection, ActivePane};

pub mod pane;
pub mod selection;

#[derive(Debug, Deserialize, Serialize, Default)]
pub(crate) struct SheetView {
    #[serde(rename = "@tabSelected", skip_serializing_if = "Option::is_none")]
    tab_selected: Option<u8>,
    #[serde(rename = "@zoomScale", skip_serializing_if = "Option::is_none")]
    zoom_scale: Option<u16>,
    #[serde(rename = "@topLeftCell", skip_serializing_if = "Option::is_none")]
    top_left_cell: Option<String>,
    #[serde(rename = "@rightToLeft", skip_serializing_if = "Option::is_none")]
    right_to_left: Option<u8>,
    #[serde(rename = "@workbookViewId")]
    workbook_view_id: u32,
    #[serde(rename = "pane", default, skip_serializing_if = "Vec::is_empty")]
    pane: Vec<Pane>,
    #[serde(rename = "selection", default, skip_serializing_if = "Vec::is_empty")]
    selection: Vec<Selection>,
}

impl SheetView {
    pub(crate) fn set_right_to_left(&mut self, right_to_left: u8) {
        self.right_to_left = Some(right_to_left);
    }

    pub(crate) fn set_tab_selected(&mut self, tab_selected: u8) {
        self.tab_selected = Some(tab_selected);
    }

    pub(crate) fn set_zoom_scale(&mut self, zoom_scale: u16) {
        self.zoom_scale = Some(zoom_scale);
    }

    pub(crate) fn set_top_left_cell(&mut self, loc_ref: &str) {
        self.top_left_cell = Some(String::from(loc_ref));
    }

    pub(crate) fn set_selection<L: LocationRange>(&mut self, loc_range: &L) {
        match self.selection.len() {
            0 => self.selection.push(Selection::from_loc_range(loc_range)),
            1 => self.selection[0].set_selection(loc_range),
            _ => {
                self.selection.iter_mut().for_each(|s| {
                    s.set_selection(loc_range)
                })
            }
        }
    }

    fn add_active_pane(&mut self, active_pane: &str) {
        let pane = Some(String::from(active_pane));
        let selection = self.selection.iter_mut()
            .find(|s| s.pane == pane);
        if let None = selection {
            self.selection.push(Selection::from_active_pane(&pane.unwrap()));
        }
    }

    pub(crate) fn set_frozen_panes<L: Location>(&mut self, loc: L, frozen: bool) {
        let (row, col) = loc.to_location();
        match (row, col) {
            (1, 1) => {},
            (1, col) => {
                self.pane = vec![Pane::from_location((1, col), "topRight")];
                self.add_active_pane("topRight");
            },
            (row, 1) => {
                self.pane = vec![Pane::from_location((row, 1), "bottomLeft")];
                self.add_active_pane("bottomLeft");
            },
            (row, col) => {
                self.pane = vec![Pane::from_location((row, col), "bottomRight")];
                self.add_active_pane("topRight");
                self.add_active_pane("bottomLeft");
                self.add_active_pane("bottomRight");
            }
        };
    }

    pub(crate) fn set_panes(&mut self, x_split: u32, y_split: u32) {
        self.pane = vec![Pane::from_split(x_split, y_split)];
        self.add_active_pane("topRight");
        self.add_active_pane("bottomLeft");
        self.add_active_pane("bottomRight");
    }
}