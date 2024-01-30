use std::cell::RefCell;
use std::cmp::Ordering;
use std::collections::BinaryHeap;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::col::Cols;
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::PhoneticPr;

#[derive(Debug, Deserialize, Serialize)]
#[serde(rename="worksheet")]
pub(crate) struct WorkSheet {
    #[serde(rename = "@xmlns")]
    xmlns: String,
    #[serde(rename = "dimension", skip_serializing_if = "Option::is_none")]
    dimension: Option<Dimension>,
    #[serde(rename = "sheetViews")]
    sheet_views: SheetViews,
    #[serde(rename = "sheetFormatPr")]
    sheet_format_pr: SheetFormatPr,
    #[serde(rename = "cols", skip_serializing_if = "Option::is_none")]
    cols: Option<Cols>,
    #[serde(rename = "sheetData")]
    pub sheet_data: SheetData,
    #[serde(rename = "phoneticPr", skip_serializing_if = "Option::is_none")]
    phonetic_pr: Option<PhoneticPr>,
    #[serde(rename = "pageMargins")]
    page_margins: PageMargins,
}

#[derive(Debug, Deserialize, Serialize)]
struct Dimension {
    #[serde(rename="@ref")]
    refer: String,
}

#[derive(Debug, Deserialize, Serialize)]
struct SheetView {
    #[serde(rename = "@tabSelected", skip_serializing_if = "Option::is_none")]
    tab_selected: Option<u32>,
    #[serde(rename = "@workbookViewId")]
    workbook_view_id: u32,
}

#[derive(Debug, Deserialize, Serialize)]
struct SheetViews {
    #[serde(rename = "sheetView")]
    sheet_view: Vec<SheetView>
}

#[derive(Debug, Deserialize, Serialize)]
struct SheetFormat {
    default_row_height: u32
}

#[derive(Debug, Deserialize, Serialize)]
struct SheetFormatPr {
    #[serde(rename = "@defaultRowHeight")]
    default_row_height: f64,
}

#[derive(Debug, Deserialize, Serialize)]
struct PageMargins {
    #[serde(rename = "@bottom")]
    bottom: f64,
    #[serde(rename = "@footer")]
    footer: f64,
    #[serde(rename = "@header")]
    header: f64,
    #[serde(rename = "@left")]
    left: f64,
    #[serde(rename = "@right")]
    right: f64,
    #[serde(rename = "@top")]
    top: f64,
}

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct SheetData {
    #[serde(rename = "row")]
    pub(crate) rows: Vec<Row>
}

impl SheetData {
    pub(crate) fn get_row(&mut self, row_id: u32) -> Option<&mut Row> {
        let row = self.rows.iter_mut().find(|r| { r.row == row_id });
        row
    }

    pub(crate) fn create_row(&mut self, row_id: u32) -> &mut Row {
        self.rows.push(Row::new(row_id));
        self.get_row(row_id).unwrap()
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Row {
    #[serde(rename = "c")]
    pub(crate) cell: Vec<Cell>,
    #[serde(rename = "@r")]
    row: u32,
}


impl Row {
    fn new(row: u32) -> Row {
        Row {
            cell: vec![],
            row
        }
    }

    pub(crate) fn get_cell(&mut self, col_id: u32) -> Option<&mut Cell> {
        let num = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0'];
        let cell = self.cell.iter_mut().find(|r| {
            let r = r.row.trim_matches(&num);
            let col = num_2_col(col_id);
            r == col
        });
        cell
    }

    pub(crate) fn create_cell(&mut self, row_id: u32, col_id: u32) -> &mut Cell {
        self.cell.push(Cell::new(row_id, col_id));
        self.cell.last_mut().unwrap()
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Cell {
    #[serde(rename = "@r")]
    row: String,
    #[serde(rename = "@s", skip_serializing_if = "Option::is_none")]
    style: Option<u32>,
    #[serde(rename = "@t")]
    text_type: String,
    #[serde(rename = "v", skip_serializing_if = "Option::is_none")]
    pub(crate) text: Option<String>,
}

impl Cell {
    fn new(row: u32, col: u32) -> Cell {
        Cell {
            row: num_2_col(col) + &row.to_string(),
            style: None,
            text_type: "s".to_string(),
            text: None,
        }
    }
}

impl WorkSheet {
    pub(crate) fn from_path<P: AsRef<Path>>(file_path: P, sheet_id: u32) -> WorkSheet {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::SheetFile(sheet_id)).unwrap();
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let work_sheet = de::from_str(&xml).unwrap();
        work_sheet
    }

    pub fn save<P: AsRef<Path>>(&self, file_path: P, sheet_id: u32) {
        let xml = se::to_string(&self).unwrap();
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::SheetFile(sheet_id)).unwrap();
        file.write_all(xml.as_ref()).unwrap();
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