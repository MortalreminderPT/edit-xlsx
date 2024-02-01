use serde::{Deserialize, Deserializer, Serialize, Serializer};
use crate::result::{RowError, RowResult};
use crate::xml::facade::EditRow;

#[derive(Debug, Serialize, Deserialize)]
pub(crate) struct SheetData {
    #[serde(rename = "row")]
    pub(crate) rows: Vec<Row>
}

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Cell {
    #[serde(flatten)]
    pub(crate) loc: CellLocation,
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
            location: num_2_col(col) + &row.to_string(),
            col: Some(col)
        }
    }

    fn col(&self) -> u32 {
        match self.col {
            Some(col) => col,
            None => col_2_num(&self.location)
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

    pub(crate) fn col(num: u32) -> u32 {
        0
    }
}

fn num_2_col(mut col_num: u32) -> String {
    let mut col = String::new();
    while col_num > 0 {
        let pop = (col_num - 1) % 26;
        col_num = (col_num - 1) / 26;
        col.push(('A' as u8 + pop as u8) as char);
    }
    col.chars().rev().collect::<String>()
}

fn col_2_num(col: &String) -> u32 {
    let num = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0'];
    let mut col = col.trim_matches(num);
    let mut col = col.chars().rev().collect::<String>();
    let mut num: u32 = 0;
    while !col.is_empty() {
        num *= 26;
        num += (col.pop().unwrap() as u8 - 'A' as u8 + 1) as u32;
    }
    num
}

#[test]
fn test_col () {
    for i in 1..1_000_000 {
        let mut s = num_2_col(i);
        let r = col_2_num(&mut s);
        assert_eq!(i, r)
    }
}