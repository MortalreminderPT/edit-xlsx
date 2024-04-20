pub(crate) mod formula;
pub(crate) mod inline_string;
mod text;

use std::fmt::Formatter;
use serde::{Deserialize, Deserializer, Serialize, Serializer};
use serde::de::{Error, Visitor};
use crate::api::cell::Cell as ApiCell;
use crate::api::cell::location::Location;
use crate::xml::worksheet::sheet_data::cell::formula::Formula;
use crate::api::cell::values::{CellDisplay, CellValue, CellType};
use crate::result::CellResult;
use crate::xml::common::FromFormat;
use crate::xml::worksheet::sheet_data::cell::inline_string::InlineString;

#[derive(Debug, Clone, Deserialize, Serialize)]
pub(crate) struct Cell {
    #[serde(rename = "@r")]
    pub(crate) loc: Sqref,
    #[serde(rename = "@s", skip_serializing_if = "Option::is_none")]
    pub(crate) style: Option<u32>,
    #[serde(rename = "@t", skip_serializing_if = "Option::is_none")]
    pub(crate) cell_type: Option<CellType>,
    #[serde(rename = "@cm", skip_serializing_if = "Option::is_none")]
    pub(crate) cell_meta_index: Option<u8>,
    #[serde(rename = "f", skip_serializing_if = "Option::is_none")]
    pub(crate) formula: Option<Formula>,
    #[serde(rename = "v", skip_serializing_if = "Option::is_none")]
    pub(crate) text: Option<String>,
    #[serde(rename = "is", skip_serializing_if = "Option::is_none")]
    pub(crate) inline_string: Option<InlineString>,
}

#[derive(Debug, Clone, Default, PartialEq)]
pub(crate) struct Sqref {
    pub(crate) col: u32,
    pub(crate) row: u32,
}

///
/// Constructor
///
impl Cell {
    pub(crate) fn new<L: Location>(loc: L) -> Cell {
        Cell {
            loc: Sqref::from_location(&loc),
            style: None,
            cell_type: Some(CellType::String),
            cell_meta_index: None,
            formula: None,
            text: None,
            inline_string: None,
        }
    }

    pub(crate) fn new_display<L: Location, T: CellDisplay + CellValue>(loc: L, text: T, style: Option<u32>) -> Cell {
        Cell {
            loc: Sqref::from_location(&loc),
            style,
            cell_type: Some(text.to_cell_type()),
            cell_meta_index: None,
            formula: None,
            text: Some(text.to_display()),
            inline_string: None,
        }
    }
}

///
/// Convertor
///
impl Cell {
    pub(crate) fn to_api_cell(&self) -> ApiCell<String> {
        let mut api_cell = ApiCell::default();
        api_cell.text = self.text.clone();
        if let Some(formula) = &self.formula {
            api_cell.formula = Some(formula.to_api_formula());
        }
        api_cell.cell_type = self.cell_type.clone();
        api_cell.style = self.style;
        if let Some(inline_string) = &self.inline_string {
            api_cell.rich_text = Some(inline_string.get_format());
        }
        api_cell
    }
}

///
/// Update
///
impl Cell {
    pub(crate) fn update_by_display<T: CellDisplay + CellValue>(&mut self, text: &T, style: Option<u32>) {
        self.text = Some(text.to_display());
        if let Some(style) = style {
            self.style = Some(style);
        }
        self.cell_type = Some(text.to_cell_type());
        self.formula = None;
    }

    pub(crate) fn update_by_api_cell<T: CellDisplay + CellValue>(&mut self, api_cell: &ApiCell<T>) -> CellResult<()> {
        if let Some(text) = &api_cell.text {
            self.text = Some(text.to_display());
            self.cell_type = api_cell.cell_type.clone();
        }
        if let Some(style) = &api_cell.style {
            self.style = Some(*style)
        }
        if let Some(formula) = &api_cell.formula {
            // if formula.formula_type.is_some() {
            //     self.text = Some(String::from("0"));
            //     self.cell_meta_index = Some(1);
            // }
            self.formula = Some(Formula::from_api_formula(formula));
        }
        if let Some(rich_text) = &api_cell.rich_text {
            self.cell_type = Some(CellType::InlineString);
            self.text = None;
            self.inline_string = Some(InlineString::from_format(rich_text));
        }
        Ok(())
    }
}

impl Sqref {
    pub(crate) fn from_location<L: Location>(location: &L) -> Sqref {
        let (row, col) = location.to_location();
        Sqref {
            col,
            row,
        }
    }

    // pub(crate) fn from_location_range<L: LocationRange>(location_range: &L) -> Sqref {
    //     let (from_row, from_col, from_row, from_col) = location_range.to_range();
    //     Sqref {
    //         col,
    //         row,
    //     }
    // }
}

///
/// Serialize and Deserialize
///
impl<'de> Visitor<'de> for Sqref {
    type Value = Sqref;

    fn expecting(&self, formatter: &mut Formatter) -> std::fmt::Result {
        todo!()
    }

    fn visit_str<E>(self, v: &str) -> Result<Self::Value, E> where E: Error {
        // let (row, col) = to_loc(&v);
        let sqref = Sqref::from_location(&v);
        Ok(sqref)
    }
}
impl<'de> Deserialize<'de> for Sqref {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error> where D: Deserializer<'de> {
        let sqref = Sqref::default();
        deserializer.deserialize_string(sqref)
    }
}
impl Serialize for Sqref {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error> where S: Serializer {
        serializer.serialize_str(&(self.row, self.col).to_ref())
    }
}
