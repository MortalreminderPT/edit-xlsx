pub(crate) mod cell;
mod row;

use serde::{Deserialize, Serialize};
use crate::api::cell::formula::FormulaType;
use crate::api::cell::location::Location;
use crate::api::cell::values::{CellDisplay, CellValue};
use crate::api::worksheet::row::RowSet;
use crate::result::RowResult;
use crate::xml::worksheet::sheet_data::cell::Cell;
use crate::xml::worksheet::sheet_data::row::Row;

#[derive(Debug, Serialize, Deserialize, Default)]
pub(crate) struct SheetData {
    #[serde(rename = "row", default)]
    rows: Vec<Row>,
}

impl SheetData {
    pub(crate) fn max_col(&self) -> u32 {
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

    pub(crate) fn set_row_by_rowset(&mut self, row: u32, row_set: &RowSet) {
        let row = self.get_or_new_row(row);
        if let Some(height) = row_set.height {
            row.height = Some(height);
            row.custom_height = Some(1);
        }
        if let Some(style) = row_set.style {
            row.style = Some(style);
            row.custom_format = Some(1);
        }
        if let Some(hidden) = row_set.hidden {
            row.hidden = Some(hidden)
        }
        if let Some(outline_level) = row_set.outline_level {
            row.outline_level = Some(outline_level)
        }
        if let Some(collapsed) = row_set.collapsed {
            row.collapsed = Some(collapsed)
        }
    }

    pub(crate) fn get_default_style(&self, row: u32) -> Option<u32> {
        let row = self.get_row(row);
        if let Some(row) = row {
            return row.style;
        }
        None
    }

    pub(crate) fn write_display<L: Location, T: CellDisplay + CellValue>(&mut self, loc: &L, text: &T, style: Option<u32>) -> RowResult<()> {
        let (row, col) = loc.to_location();
        let row = self.get_or_new_row(row);
        row.add_display_cell(col, text, style);
        Ok(())
    }

    pub(crate) fn write_formula<L: Location>(&mut self, loc: &L, formula: &str, formula_type: FormulaType, style: Option<u32>) -> RowResult<()> {
        let (row, col) = loc.to_location();
        let row = self.get_or_new_row(row);
        row.add_formula_cell(col, formula, formula_type, style);
        Ok(())
    }
}

trait _OrderRow {
    fn get_position_by_row(&self, row: u32) -> usize;
    fn new_row(&mut self, row: u32) -> &mut Row;
    fn get_row_mut(&mut self, row: u32) -> Option<&mut Row>;
    fn get_row(&self, row: u32) -> Option<&Row>;
    fn get_or_new_row(&mut self, row: u32) -> &mut Row;
    fn get_last_row(&self) -> Option<&Row>;
}

impl _OrderRow for SheetData {
    fn get_position_by_row(&self, row: u32) -> usize {
        let mut l = 0;
        let mut r = self.rows.len();
        while r - l > 0 {
            let mid = (l + r) / 2;
            if row == self.rows[mid].row {
                return mid;
            }
            else if row < self.rows[mid].row {
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
    fn get_row_mut(&mut self, row: u32) -> Option<&mut Row> {
        let r = self.get_position_by_row(row);
        if r >= self.rows.len() {return None}
        return if row == self.rows[r].row { Some(&mut self.rows[r]) } else { None }
    }

    fn get_row(&self, row: u32) -> Option<&Row> {
        let r = self.get_position_by_row(row);
        if r >= self.rows.len() {return None}
        return if row == self.rows[r].row { Some(&self.rows[r]) } else { None }
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
    fn get_last_row(&self) -> Option<&Row> {
        self.rows.last()
    }
}
