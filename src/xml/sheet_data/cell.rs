mod location;
pub(crate) mod formula;
pub(crate) mod values;

use std::hash::{Hash, Hasher};
use serde::{Deserialize, Serialize};
use crate::xml::sheet_data::cell::formula::{Formula, FormulaType};
use crate::xml::sheet_data::cell::location::Location;
use crate::xml::sheet_data::cell::values::{CellDisplay, CellValue, CellType};

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Cell {
    #[serde(flatten)]
    pub(crate) loc: Location,
    #[serde(rename = "@s", skip_serializing_if = "Option::is_none")]
    pub(crate) style: Option<u32>,
    #[serde(rename = "@t", default = "CellType::default", serialize_with = "CellType::se", deserialize_with = "CellType::de")]
    pub(crate) cell_type: CellType,
    #[serde(rename = "f", skip_serializing_if = "Option::is_none")]
    pub(crate) formula: Option<Formula>,
    #[serde(rename = "v", skip_serializing_if = "Option::is_none")]
    pub(crate) text: Option<String>,
}

impl Cell {
    pub(crate) fn new_display<T: CellDisplay + CellValue>(row: u32, col: u32, text: T, style: Option<u32>) -> Cell {
        Cell {
            loc: Location::from_row_and_col(row, col),
            style,
            cell_type: text.to_cell_type(),
            formula: None,
            text: Some(text.to_display()),
        }
    }

    pub(crate) fn new_formula(row: u32, col: u32, formula: &str, formula_type: FormulaType, style: Option<u32>) -> Cell {
        let loc = Location::from_row_and_col(row, col);
        let formula = Formula::from_formula_type(formula, formula_type);
        Cell {
            loc,
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

impl Cell {
    pub(crate) fn new<T: CellDisplay + CellValue>(row: u32, col: u32, text: T, formula: Option<Formula>, style_id: Option<u32>) -> Cell {
        Cell {
            loc: Location::from_row_and_col(row, col),
            style: style_id,
            cell_type: text.to_cell_type(),
            formula,
            text: Some(text.to_display()),
        }
    }

    ///
    /// 更新当前Cell中的内容，包括text及style
    /// # Arguments
    ///
    /// * `text`: 更新的文字内容，必须是同时实现CellDisplay和CellType
    /// * `style_id`: 更新的style id
    ///
    pub(crate) fn update_value<T: CellDisplay + CellValue>(&mut self, text: T, formula: Option<Formula>, style_id: Option<u32>) {
        self.cell_type = text.to_cell_type();
        self.text = Some(text.to_display());
        self.formula = formula;
        self.style = style_id;
    }
}
