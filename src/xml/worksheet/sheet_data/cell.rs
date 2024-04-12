pub(crate) mod formula;

use std::fmt::Formatter;
use serde::{Deserialize, Deserializer, Serialize, Serializer};
use serde::de::{Error, Visitor};
use crate::api::cell::formula::FormulaType;
use crate::api::cell::location::{Location, LocationRange};
use crate::xml::worksheet::sheet_data::cell::formula::Formula;
use crate::api::cell::values::{CellDisplay, CellValue, CellType};

#[derive(Debug, Clone, Deserialize, Serialize)]
pub(crate) struct Cell {
    #[serde(rename = "@r")]
    pub(crate) loc: Sqref,
    #[serde(rename = "@s", skip_serializing_if = "Option::is_none")]
    pub(crate) style: Option<u32>,
    #[serde(rename = "@t", skip_serializing_if = "Option::is_none")]
    pub(crate) cell_type: Option<CellType>,
    #[serde(rename = "@cm", skip_serializing_if = "Option::is_none")]
    pub(crate) cm: Option<u8>,
    #[serde(rename = "f", skip_serializing_if = "Option::is_none")]
    pub(crate) formula: Option<Formula>,
    #[serde(rename = "v", skip_serializing_if = "Option::is_none")]
    pub(crate) text: Option<String>,
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
            cm: None,
            formula: None,
            text: None,
        }
    }

    pub(crate) fn new_display<L: Location, T: CellDisplay + CellValue>(loc: L, text: T, style: Option<u32>) -> Cell {
        Cell {
            loc: Sqref::from_location(&loc),
            style,
            cell_type: Some(text.to_cell_type()),
            cm: None,
            formula: None,
            text: Some(text.to_display()),
        }
    }

    pub(crate) fn new_formula<L: Location>(loc: L, formula: &str, formula_type: FormulaType, style: Option<u32>) -> Cell {
        let formula = Formula::from_formula_type(formula, formula_type);
        Cell {
            loc: Sqref::from_location(&loc),
            style,
            cell_type: Some(CellType::String),
            cm: Some(1),
            formula: Some(formula),
            text: None,
        }
    }
}

///
/// Get
///
impl Cell {}

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

    pub(crate) fn update<T: CellDisplay + CellValue>(
        &mut self,
        text: Option<&T>,
        formula: Option<&str>,
        formula_type: Option<FormulaType>,
        style: Option<u32>
    ) {
        if let Some(text) = text {
            self.update_by_display(text, style);
        }
        if let (Some(formula), Some(formula_type)) = (formula, formula_type) {
            self.update_by_formula(formula, formula_type, style)
        }
    }

    pub(crate) fn update_by_formula(&mut self, formula: &str, formula_type: FormulaType, style: Option<u32>) {
        self.style = style;
        self.cell_type = None;
        match formula_type {
            FormulaType::OldFormula(_) => {
                self.text = None;
                self.cm = None;
            },
            _ => {
                self.text = Some(String::from("0"));
                self.cm = Some(1);
            },
        };
        let formula = Formula::from_formula_type(formula, formula_type);
        self.formula = Some(formula);
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
        // let sqref = Sqref::from_location(&v);
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
