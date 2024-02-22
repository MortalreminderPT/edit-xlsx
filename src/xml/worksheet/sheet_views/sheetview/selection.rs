use serde::{Deserialize, Serialize};
use crate::api::cell::location::{Location, LocationRange};
use crate::xml::worksheet::sheet_data::cell::Sqref;

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Selection {
    #[serde(rename = "@pane", skip_serializing_if = "Option::is_none")]
    pub(crate) pane: Option<String>,
    #[serde(rename = "@activeCell", skip_serializing_if = "Option::is_none")]
    active_cell: Option<Sqref>,
    #[serde(rename = "@sqref", skip_serializing_if = "Option::is_none")]
    sqref: Option<String>,
}

impl Default for Selection {
    fn default() -> Self {
        Self {
            active_cell: Some(Sqref::default()),
            sqref: None,
            pane: None,
        }
    }
}

impl Selection {
    pub(crate) fn set_selection<L: LocationRange>(&mut self, loc_range: &L) {
        // let mut sqref = self.sqref.take().unwrap_or_default();
        // sqref = String::from(loc_ref);
        self.active_cell = Some(Sqref::from_location(&loc_range.end_ref().as_str()));
        self.sqref = Some(loc_range.to_range_ref());
    }

    pub(crate) fn from_loc_range<L: LocationRange>(loc_range: &L) -> Selection {
        let mut selection = Selection::default();
        selection.set_selection(loc_range);
        selection
    }

    pub(crate) fn from_active_pane(pane: &str) -> Selection {
        let mut selection = Selection::default();
        selection.pane = Some(pane.to_string());
        selection
    }

    pub(crate) fn default_pane<L: Location>(selection_pane: ActivePane<L>) -> Self {
        Self {
            pane: Some(String::from(selection_pane.get_pane())),
            active_cell: selection_pane.get_active_cell(),
            sqref: selection_pane.get_sqref(),
        }
    }

    pub(crate) fn by_pane(pane: &str) -> Self {
        Self {
            pane: Some(pane.to_string()),
            active_cell: None,
            sqref: None,
        }
    }

    pub(crate) fn update_by_pane(&mut self, selection_pane: &str) {
        self.pane = Some(String::from(selection_pane));
    }
}

pub(crate) enum ActivePane<L: Location> {
    TopRight(Option<L>),
    BottomLeft(Option<L>),
    BottomRight(Option<L>),
}

impl<L: Location> ActivePane<L> {
    pub(crate) fn get_pane(&self) -> &str {
        match self {
            ActivePane::TopRight(_) => "topRight",
            ActivePane::BottomLeft(_) => "bottomLeft",
            ActivePane::BottomRight(_) => "bottomRight",
        }
    }

    pub(crate) fn get_active_cell(&self) -> Option<Sqref> {
        match self {
            ActivePane::TopRight(l) => match l {
                Some(l) => Some(Sqref::from_location(l)),
                None => None,
            },
            ActivePane::BottomLeft(l) => match l {
                Some(l) => Some(Sqref::from_location(l)),
                None => None,
            },
            ActivePane::BottomRight(l) => match l {
                Some(l) => Some(Sqref::from_location(l)),
                None => None,
            },
        }
    }

    pub(crate) fn get_sqref(&self) -> Option<String> {
        match self {
            ActivePane::TopRight(l) => match l {
                Some(l) => Some(l.to_ref()),
                None => None,
            },
            ActivePane::BottomLeft(l) => match l {
                Some(l) => Some(l.to_ref()),
                None => None,
            },
            ActivePane::BottomRight(l) => match l {
                Some(l) => Some(l.to_ref()),
                None => None,
            },
        }
    }

    pub(crate) fn get_split(&self) -> (Option<u32>, Option<u32>) {
        let (x_split, y_split) = match self {
            ActivePane::TopRight(l) => match l {
                Some(l) => {
                    let (_, col) = l.to_location();
                    (None, Some(col - 1))
                },
                None => (None, None),
            },
            ActivePane::BottomLeft(l) => match l {
                Some(l) => {
                    let (row, _) = l.to_location();
                    (Some(row - 1), None)
                },
                None => (None, None),
            },
            ActivePane::BottomRight(l) => match l {
                Some(l) => {
                    let (row, col) = l.to_location();
                    (Some(row - 1), Some(col - 1))
                },
                None => (None, None),
            },
        };
        (x_split, y_split)
    }
}