use std::hash::{Hash, Hasher};
use serde::{Deserialize, Serialize};
use crate::utils::col_helper;

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Location {
    #[serde(rename = "@r")]
    location_ref: String,
    #[serde(skip)]
    row: Option<u32>,
    #[serde(skip)]
    col: Option<u32>,
}

impl Location {
    pub(crate) fn from_ref(location_ref: &str) -> Location {
        let (row, col) = col_helper::to_loc(location_ref);
        Location {
            location_ref: String::from(location_ref),
            row: Some(row),
            col: Some(col),
        }
    }

    pub(crate) fn from_row_and_col(row: u32, col: u32) -> Location {
        Location {
            location_ref: col_helper::to_ref(row, col),
            row: Some(row),
            col: Some(col),
        }
    }
    
    pub(crate) fn location_ref(&self) -> &str {
        &self.location_ref
    }

    pub(crate) fn row(&self) -> u32 {
        if let None = self.row {
            let (row, col) = col_helper::to_loc(&self.location_ref);
            // self.row = Some(row);
            // self.col = Some(col);
            return row;
        }
        self.row.unwrap()
    }

    pub(crate) fn col(&self) -> u32 {
        if let None = self.col {
            let (row, col) = col_helper::to_loc(&self.location_ref);
            // self.row = Some(row);
            // self.col = Some(col);
            return col;
        }
        self.col.unwrap()
    }
}