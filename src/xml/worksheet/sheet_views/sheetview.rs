use serde::{Deserialize, Serialize};
use crate::api::cell::location::Location;
use crate::xml::worksheet::sheet_views::sheetview::pane::Pane;
use crate::xml::worksheet::sheet_views::sheetview::selection::{Selection, ActivePane};

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
    #[serde(rename = "pane", default, skip_serializing_if = "Vec::is_empty")]
    pane: Vec<Pane>,
    #[serde(rename = "selection", default, skip_serializing_if = "Vec::is_empty")]
    selection: Vec<Selection>,
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
        if self.selection.len() == 0 {
            self.selection.push(Selection::default());
        }
        self.selection.last_mut().unwrap().set_selection(&loc_ref);
    }

    fn update_by_pane<L: Location>(&mut self, active_pane: ActivePane<L>) {
        if self.selection.len() == 0 {
            self.selection.push(Selection::default());
        }
        self.selection.last_mut().unwrap().update_by_pane(active_pane);
    }

    pub(crate) fn set_freeze_panes<L: Location>(&mut self, loc: L) {
        let (row, col) = loc.to_location();
        match (row, col) {
            (1, 1) => {},
            (1, col) => {
                self.pane = vec![Pane::default_pane(ActivePane::TopRight(Some((1, col))))];
                self.update_by_pane(ActivePane::<L>::TopRight(None));
            },
            (row, 1) => {
                self.pane = vec![Pane::default_pane(ActivePane::BottomLeft(Some((row, 1))))];
                self.update_by_pane(ActivePane::<L>::BottomLeft(None));
            },
            (row, col) => {
                self.pane = vec![Pane::default_pane(ActivePane::BottomRight(Some((row, col))))];
                self.update_by_pane(ActivePane::<L>::BottomRight(None));
                self.selection.append(&mut vec![
                    Selection::default_pane(ActivePane::TopRight(Some((1, col)))),
                    Selection::default_pane(ActivePane::BottomLeft(Some((row, 1)))),
                ]);
            }
        };
    }
}