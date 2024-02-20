pub(crate) mod formula;

use std::fmt::Formatter;
use serde::{Deserialize, Deserializer, Serialize, Serializer};
use serde::de::{Error, Visitor};
use crate::api::cell::formula::FormulaType;
use crate::api::cell::location::Location;
use crate::xml::worksheet::sheet_data::cell::formula::Formula;
use crate::api::cell::values::{CellDisplay, CellValue, CellType};

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Cell {
    #[serde(rename = "@r")]
    pub(crate) loc: Sqref,
    #[serde(rename = "@s", skip_serializing_if = "Option::is_none")]
    pub(crate) style: Option<u32>,
    #[serde(rename = "@t", default = "CellType::default", serialize_with = "CellType::se", deserialize_with = "CellType::de")]
    pub(crate) cell_type: CellType,
    #[serde(rename = "f", skip_serializing_if = "Option::is_none")]
    pub(crate) formula: Option<Formula>,
    #[serde(rename = "v", skip_serializing_if = "Option::is_none")]
    pub(crate) text: Option<String>,
}

#[derive(Debug, Default)]
pub(crate) struct Sqref {
    pub(crate) col: u32,
    pub(crate) row: u32,
}

impl Sqref {
    pub(crate) fn from_location<L: Location>(location: &L) -> Sqref {
        let (row, col) = location.to_location();
        Sqref {
            col,
            row,
        }
    }
}

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

impl Cell {
    pub(crate) fn new<L: Location>(loc: L) -> Cell {
        Cell {
            loc: Sqref::from_location(&loc),
            style: None,
            cell_type: CellType::String,
            formula: None,
            text: None,
        }
    }

    pub(crate) fn new_display<L: Location, T: CellDisplay + CellValue>(loc: L, text: T, style: Option<u32>) -> Cell {
        Cell {
            loc: Sqref::from_location(&loc),
            style,
            cell_type: text.to_cell_type(),
            formula: None,
            text: Some(text.to_display()),
        }
    }

    pub(crate) fn new_formula<L: Location>(loc: L, formula: &str, formula_type: FormulaType, style: Option<u32>) -> Cell {
        let formula = Formula::from_formula_type(formula, formula_type);
        Cell {
            loc: Sqref::from_location(&loc),
            style,
            cell_type: CellType::String,
            formula: Some(formula),
            text: None,
        }
    }
}

impl Cell {
    pub(crate) fn update_by_display<T: CellDisplay + CellValue>(&mut self, text: T, style: Option<u32>) {
        self.text = Some(text.to_display());
        if let Some(style) = style {
            self.style = Some(style);
        }
        self.cell_type = text.to_cell_type();
        self.formula = None;
    }

    pub(crate) fn update_by_formula(&mut self, formula: &str, formula_type: FormulaType, style: Option<u32>) {
        let formula = Formula::from_formula_type(formula, formula_type);
        self.formula = Some(formula);
        self.style = style;
        self.cell_type = CellType::String;
        self.text = None;
    }
}