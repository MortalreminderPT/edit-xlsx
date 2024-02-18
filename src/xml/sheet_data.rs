pub(crate) mod cell;
mod row;

use std::cmp::max;
use serde::{Deserialize, Serialize};
use crate::result::{RowError, RowResult};
pub(crate) use crate::xml::sheet_data::cell::Cell;
use crate::xml::sheet_data::cell::formula::FormulaType;
use crate::xml::sheet_data::cell::values::{CellDisplay, CellValue};
use crate::xml::sheet_data::row::Row;

#[derive(Debug, Serialize, Deserialize)]
pub(crate) struct SheetData {
    #[serde(rename = "row", default = "Vec::new")]
    pub(crate) rows: Vec<Row>,
    #[serde(skip)]
    max_row: u32,
    #[serde(skip)]
    max_col: Option<u32>,
}

impl Default for SheetData {
    fn default() -> SheetData {
        SheetData {
            rows: vec![],
            max_row: 0,
            max_col: None,
        }
    }
}

impl SheetData {
    pub(crate) fn get_row(&mut self, row: u32) -> RowResult<&mut Row> {
        Ok(self.rows
            .iter_mut()
            .find(|r| { r.row == row })
            .ok_or(RowError::RowNotFound)?)
    }

    pub(crate) fn add_row(&mut self, row: u32) -> &mut Row {
        self.rows.push(Row::new(row));
        self.max_row = max(self.max_row, row);
        self.rows.last_mut().unwrap()
    }

    pub(crate) fn set_row(&mut self, row: u32, height: f64, style: Option<u32>) -> RowResult<()> {
        let sheet_data_row = match self.get_row(row) {
            Ok(row) => row,
            Err(_) => self.add_row(row)
        };
        sheet_data_row.height = height;
        if let None = sheet_data_row.custom_height {
            sheet_data_row.custom_height = Some(1);
        }
        if let Some(style) = style {
            sheet_data_row.style = Some(style);
            if let None = sheet_data_row.custom_format {
                sheet_data_row.custom_format = Some(1);
            }
        }
        Ok(())
    }

    pub(crate) fn max_col(&mut self) -> u32 {
        match self.max_col {
            Some(col) => col,
            None => {
                let col = self.rows.iter_mut().max_by_key(|r| { r.max_col() });
                match col {
                    Some(col) => col.max_col(),
                    None => 0,
                }
            }
        }
    }

    pub(crate) fn sort_row(&mut self) {
        self.rows.iter_mut().for_each(|r| r.cells.sort_by_key(|c: &Cell| c.loc.col()));
        self.rows.sort_by_key(|r| r.row)
    }

    pub(crate) fn write_display<T: CellDisplay + CellValue>(&mut self, row: u32, col: u32, text: T, style: Option<u32>) {
        let sheet_data_row = match self.get_row(row) {
            Ok(row) => row,
            Err(_) => self.add_row(row)
        };
        sheet_data_row.add_display_cell(col, text, style);
    }

    pub(crate) fn write_formula(&mut self, row: u32, col: u32, formula: &str, formula_type: FormulaType, style: Option<u32>) {
        let sheet_data_row = match self.get_row(row) {
            Ok(row) => row,
            Err(_) => self.add_row(row)
        };
        sheet_data_row.add_formula_cell(col, formula, formula_type, style);
    }
}
