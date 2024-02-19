use std::cmp::max;
use serde::{Deserialize, Serialize};
use crate::xml::sheet_data::Cell;
use crate::xml::sheet_data::cell::formula::FormulaType;
use crate::xml::sheet_data::cell::values::{CellDisplay, CellValue};

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Row {
    #[serde(rename = "c")]
    pub(crate) cells: Vec<Cell>,
    #[serde(rename = "@r")]
    pub(crate) row: u32,
    #[serde(rename = "@ht")]
    pub(crate) height: Option<f64>,
    #[serde(rename = "@s", skip_serializing_if = "Option::is_none")]
    pub(crate) style: Option<u32>,
    #[serde(rename = "@customFormat", skip_serializing_if = "Option::is_none")]
    pub(crate) custom_format: Option<u32>,
    #[serde(rename = "@customHeight", skip_serializing_if = "Option::is_none")]
    pub(crate) custom_height: Option<u32>,
    #[serde(skip)]
    max_col: Option<u32>,
}


impl Row {
    pub(crate) fn new(row: u32) -> Row {
        Row {
            cells: vec![],
            row,
            height: None,
            style: None,
            custom_format: None,
            custom_height: None,
            max_col: None,
        }
    }

    pub(crate) fn get_cell(&mut self, col_id: u32) -> Option<&mut Cell> {
        let cell = self.cells
            .iter_mut()
            .find(|r| {
                r.loc.col() == col_id
            });
        cell
    }

    pub(crate) fn max_col(&self) -> u32 {
        match self.max_col {
            Some(col) => col,
            None => self.cells.iter().max_by_key(|&c| { c.loc.col() }).unwrap().loc.col(),
        }
    }

    pub(crate) fn add_display_cell<T: CellDisplay + CellValue>(&mut self, col: u32, text: T, style: Option<u32>) {
        // 判断新增cell位置是否已经存在别的cell
        let cell = self.cells.iter_mut().find(|c| { c.loc.col() == col });
        match cell {
            Some(cell) => cell.update_by_display(text, style),
            None => {
                let cell = Cell::new_display(self.row, col, text, style);
                self.max_col = max(self.max_col, Some(col));
                self.cells.push(cell);
            }
        }
    }

    pub(crate) fn add_formula_cell(&mut self, col: u32, formula: &str, formula_type: FormulaType, style: Option<u32>) {
        let other_cell = self.cells.iter_mut().find(|c| { c.loc.col() == col });
        match other_cell {
            Some(cell) => cell.update_by_formula(formula, formula_type, style),
            None => {
                let cell = Cell::new_formula(self.row, col, formula, formula_type, style);
                self.max_col = max(self.max_col, Some(col));
                self.cells.push(cell);
            }
        };
    }

    fn add_url_cell(&mut self, col: u32, text: &str, style: Option<u32>) {
    }
}