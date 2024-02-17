pub(crate) mod cell;
pub(crate) mod cell_values;

use std::cmp::max;
use serde::{Deserialize, Serialize};
use crate::result::{RowError, RowResult};
pub(crate) use crate::xml::sheet_data::cell::Cell;
use crate::xml::sheet_data::cell_values::{CellDisplay, CellType};

#[derive(Debug, Serialize, Deserialize)]
pub(crate) struct SheetData {
    #[serde(rename = "row", default = "Vec::new")]
    pub(crate) rows: Vec<Row>,
    #[serde(skip)]
    pub(crate) max_row_id: u32,
    #[serde(skip)]
    pub(crate) max_col_id: Option<u32>,
}

impl Default for SheetData {
    fn default() -> SheetData {
        SheetData {
            rows: vec![],
            max_row_id: 0,
            max_col_id: None,
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Row {
    #[serde(rename = "c")]
    pub(crate) cells: Vec<Cell>,
    #[serde(rename = "@r")]
    pub(crate) row: u32,
    #[serde(rename = "@ht")]
    pub(crate) height: f64,
    #[serde(rename = "@s", skip_serializing_if = "Option::is_none")]
    pub(crate) style: Option<u32>,
    #[serde(rename = "@customFormat", skip_serializing_if = "Option::is_none")]
    custom_format: Option<u32>,
    #[serde(rename = "@customHeight", skip_serializing_if = "Option::is_none")]
    custom_height: Option<u32>,
    #[serde(skip)]
    max_col_id: Option<u32>,
}

impl SheetData {
    pub(crate) fn get_row(&mut self, row_id: u32) -> RowResult<&mut Row> {
        let row = self.rows
            .iter_mut()
            .find(|r| { r.row == row_id })
            .ok_or(RowError::RowNotFound)?;
        Ok(row)
    }

    pub(crate) fn create_row(&mut self, row_id: u32) -> &mut Row {
        self.rows.push(Row::new(row_id));
        self.max_row_id = max(self.max_row_id, row_id);
        // self.get_row(row_id).unwrap()
        self.rows.last_mut().unwrap()
    }

    pub(crate) fn update_or_create_row(&mut self, row_id: u32, height: f64, style_id: Option<u32>) -> RowResult<&mut Row> {
        match self.get_row(row_id) {
            Ok(row) => row,
            Err(_) => self.create_row(row_id),
        };
        let row = self.get_row(row_id).unwrap();
        row.height = height;
        if let None = row.custom_height {
            row.custom_height = Some(1);
        }
        if let Some(style_id) = style_id {
            row.style = Some(style_id);
            if let None = row.custom_format {
                row.custom_format = Some(1);
            }
        }
        Ok(row)
    }

    fn update_row(&mut self, row_id: u32) -> RowResult<&mut Row> {
        todo!()
    }

    fn delete_row(&mut self, row_id: u32) -> RowResult<()> {
        todo!()
    }

    pub(crate) fn max_col(&mut self) -> u32 {
        match self.max_col_id {
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
}

impl Row {
    pub(crate) fn new(row: u32) -> Row {
        Row {
            cells: vec![],
            row,
            height: 15.0,
            style: None,
            custom_format: None,
            custom_height: None,
            max_col_id: None,
        }
    }

    pub(crate) fn get_cell(&mut self, col_id: u32) -> Option<&mut Cell> {
        let cell = self.cells
            .iter_mut()
            .find(|mut r| {
                r.loc.col() == col_id
            });
        cell
    }

    pub(crate) fn create_cell<T: CellDisplay + CellType>(&mut self, col_id: u32, text: T, style_id: Option<u32>) -> &mut Cell {
        self.cells.push(Cell::new(self.row, col_id, text, style_id));
        self.max_col_id = max(self.max_col_id, Some(col_id));
        self.cells.last_mut().unwrap()
    }

    pub(crate) fn max_col(&self) -> u32 {
        match self.max_col_id {
            Some(col) => col,
            None => self.cells.iter().max_by_key(|&c| { c.loc.col() }).unwrap().loc.col(),
        }
    }

    pub(crate) fn update_or_create_cell<T: CellDisplay + CellType>(&mut self, col_id: u32, text: T, style_id: Option<u32>) {
        let cell = self.get_cell(col_id);
        match cell {
            Some(cell) => {
                cell.update_value(text, style_id);
            }
            None => {
                self.create_cell(col_id, text, style_id);
            }
        }
    }
}