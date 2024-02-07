pub(crate) mod cell;
pub(crate) mod cell_values;

use serde::{Deserialize, Serialize};
use crate::result::{RowError, RowResult};
use crate::xml::manage::EditRow;
pub(crate) use crate::xml::sheet_data::cell::Cell;
use crate::xml::sheet_data::cell_values::{CellDisplay, CellType};

#[derive(Debug, Serialize, Deserialize)]
pub(crate) struct SheetData {
    #[serde(rename = "row", default = "Vec::new")]
    pub(crate) rows: Vec<Row>
}

impl SheetData {
    pub(crate) fn default() -> SheetData {
        SheetData {
            rows: vec![]
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Row {
    #[serde(rename = "c")]
    pub(crate) cells: Vec<Cell>,
    #[serde(rename = "@r")]
    pub(crate) row: u32,
}

impl EditRow for SheetData {
    fn get(&mut self, row_id: u32) -> RowResult<&mut Row> {
        let row = self.rows
            .iter_mut()
            .find(|r| { r.row == row_id })
            .ok_or(RowError::RowNotFound)?;
        Ok(row)
    }

    fn create(&mut self, row_id: u32) -> RowResult<&mut Row> {
        self.rows.push(Row::new(row_id));
        self.get(row_id)
    }

    fn update(&mut self, row_id: u32) -> RowResult<&mut Row> {
        todo!()
    }

    fn delete(&mut self, row_id: u32) -> RowResult<()> {
        todo!()
    }

    fn sort(&mut self) {
        self.rows.iter_mut().for_each(|r| r.cells.sort_by_key(|c| c.loc.col()));
        self.rows.sort_by_key(|r| r.row)
    }
}

impl Row {
    pub(crate) fn new(row: u32) -> Row {
        Row {
            cells: vec![],
            row
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

    pub(crate) fn create_cell<T: CellDisplay + CellType>(&mut self, col_id: u32, text: T, style_id: Option<u32>) -> &mut Cell {
        self.cells.push(Cell::new(self.row, col_id, text, style_id));
        self.cells.last_mut().unwrap()
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
