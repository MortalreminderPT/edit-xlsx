pub(crate) mod cell;
mod row;

use serde::{Deserialize, Serialize};
use crate::api::cell::formula::FormulaType;
use crate::api::cell::values::{CellDisplay, CellValue};
use crate::result::RowResult;
use crate::xml::worksheet::sheetdata::cell::Cell;
use crate::xml::worksheet::sheetdata::row::Row;

#[derive(Debug, Serialize, Deserialize, Default)]
pub(crate) struct SheetData {
    #[serde(rename = "row", default)]
    rows: Vec<Row>,
}

impl SheetData {
    fn get_position_by_row(&self, row: u32) -> usize {
        let mut l = 0;
        let mut r = self.rows.len();
        while r - l > 0 {
            let mid = (l + r) / 2;
            if row < self.rows[mid].row {
                r = mid;
            } else {
                l = mid + 1;
            }
        }
        r
    }

    fn new_row(&mut self, row: u32) -> &mut Row {
        let r = self.get_position_by_row(row);
        self.rows.insert(r, Row::new(row));
        return &mut self.rows[r];
    }

    fn get_row(&mut self, row: u32) -> Option<&mut Row> {
        let r = self.get_position_by_row(row);
        return if row == self.rows[r].row { Some(&mut self.rows[r]) } else { None }
    }

    fn get_or_new_row(&mut self, row: u32) -> &mut Row {
        let r = self.get_position_by_row(row);
        return if r < self.rows.len() && self.rows[r].row == row {
            &mut self.rows[r]
        } else {
            self.rows.insert(r, Row::new(row));
            &mut self.rows[r]
        }
    }

    pub(crate) fn set_row(&mut self, row: u32, height: f64, style: Option<u32>) -> RowResult<()> {
        let sheet_data_row = self.get_or_new_row(row);
        sheet_data_row.height = Some(height);
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
        match self.rows.iter().max_by_key(|row| { row.max_col() }) {
            Some(row) => row.max_col(),
            None => 0
        }
    }

    pub(crate) fn max_row(&self) -> u32 {
        match self.rows.last() {
            Some(row) => row.row,
            None => 0
        }
    }

    pub(crate) fn sort_row(&mut self) {
        // self.rows.iter_mut().for_each(|r| r.cells.sort_by_key(|c: &Cell| c.loc.col));
        self.rows.sort_by_key(|r| r.row)
    }

    pub(crate) fn write_display<T: CellDisplay + CellValue>(&mut self, row: u32, col: u32, text: T, style: Option<u32>) {
        // let sheet_data_row = match self.get_row(row) {
        //     Ok(row) => row,
        //     Err(_) => self.add_row(row)
        // };
        let sheet_data_row = self.get_or_new_row(row);
        sheet_data_row.add_display_cell(col, text, style);
    }

    pub(crate) fn write_formula(&mut self, row: u32, col: u32, formula: &str, formula_type: FormulaType, style: Option<u32>) {
        // let sheet_data_row = match self.get_row(row) {
        //     Ok(row) => row,
        //     Err(_) => self.add_row(row)
        // };
        let sheet_data_row = self.get_or_new_row(row);
        sheet_data_row.add_formula_cell(col, formula, formula_type, style);
    }
}
