use serde::{Deserialize, Serialize};
use crate::utils::col_helper;
use crate::xml::sheet_data::cell_values::{CellDisplay, CellType, CellValues};

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Cell {
    #[serde(flatten)]
    pub(crate) loc: CellLocation,
    #[serde(rename = "@s", skip_serializing_if = "Option::is_none")]
    pub(crate) style: Option<u32>,
    #[serde(rename = "@t", default = "CellValues::default", serialize_with = "CellValues::se", deserialize_with = "CellValues::de")]
    pub(crate) cell_values: CellValues,
    #[serde(rename = "f", skip_serializing_if = "Option::is_none")]
    pub(crate) formula: Option<Formula>,
    #[serde(rename = "v", skip_serializing_if = "Option::is_none")]
    pub(crate) text: Option<String>,
}

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Formula {
    #[serde(rename = "@t", skip_serializing_if = "Option::is_none")]
    formula_type: Option<String>,
    #[serde(rename = "@ref", skip_serializing_if = "Option::is_none")]
    formula_ref: Option<String>,
    #[serde(rename = "$value")]
    formula: String,
}

impl Formula {
    pub(crate) fn from_formula(formula: String) -> Formula {
        Formula {
            formula_type: None,
            formula_ref: None,
            formula,
        }
    }

    pub(crate) fn from_array_formula(formula: String, row: u32, col: u32) -> Formula {
        Formula {
            formula_type: Some("array".to_string()),
            formula_ref: Some(col_helper::to_ref(row, col)),
            formula,
        }
    }

    pub(crate) fn from_ref_formula(formula: String, formula_ref: String) -> Formula {
        Formula {
            formula_type: Some("array".to_string()),
            formula_ref: Some(formula_ref),
            formula,
        }
    }
}

impl Cell {
    pub(crate) fn new<T: CellDisplay + CellType>(row: u32, col: u32, text: T, formula: Option<Formula>, style_id: Option<u32>) -> Cell {
        let loc = CellLocation::new(row, col);
        Cell {
            loc,
            style: style_id,
            cell_values: text.to_cell_val(),
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
    pub(crate) fn update_value<T: CellDisplay + CellType>(&mut self, text: T, formula: Option<Formula>, style_id: Option<u32>) {
        self.cell_values = text.to_cell_val();
        self.text = Some(text.to_display());
        self.formula = formula;
        self.style = style_id;
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct CellLocation {
    #[serde(rename = "@r")]
    location: String,
    #[serde(skip)]
    col: Option<u32>,
}

impl CellLocation {
    fn new(row: u32, col: u32) -> CellLocation {
        let location = col_helper::to_col_name(col) + &row.to_string();
        CellLocation {
            col: Some(col_helper::to_col(&location)),
            location,
        }
    }

    pub(crate) fn col(&self) -> u32 {
        match self.col {
            Some(col) => col,
            None => col_helper::to_col(&self.location)
        }
    }
}
