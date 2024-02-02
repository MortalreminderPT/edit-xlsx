use serde::{Deserialize, Serialize};
use crate::result::{RowError, RowResult};
use crate::utils::col_helper;
use crate::utils::col_helper::to_col_name;
use crate::xml::facade::EditRow;

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
pub(crate) struct Cell {
    #[serde(flatten)]
    loc: CellLocation,
    #[serde(rename = "@s", skip_serializing_if = "Option::is_none")]
    pub(crate) style: Option<u32>,
    #[serde(rename = "@t")]
    pub(crate) text_type: String,
    #[serde(rename = "v", skip_serializing_if = "Option::is_none")]
    pub(crate) text: Option<String>,
}

#[derive(Debug, Deserialize, Serialize)]
struct CellLocation {
    #[serde(rename = "@r")]
    location: String,
    #[serde(skip)]
    col: Option<u32>,
}

impl CellLocation {
    fn new(row: u32, col: u32) -> CellLocation {
        CellLocation {
            location: to_col_name(col) + &row.to_string(),
            col: Some(col)
        }
    }

    fn col(&self) -> u32 {
        match self.col {
            Some(col) => col,
            None => col_helper::to_col(&self.location)
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

    pub(crate) fn create_cell(&mut self, row_id: u32, col_id: u32) -> &mut Cell {
        self.cells.push(Cell::new(row_id, col_id));
        self.cells.last_mut().unwrap()
    }
}


impl Cell {
    fn new(row: u32, col: u32) -> Cell {
        Cell {
            loc: CellLocation::new(row, col), // num_2_col(col) + &row.to_string(),
            style: None,
            text_type: "s".to_string(),
            text: None,
        }
    }
}