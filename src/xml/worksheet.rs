use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::common::{PhoneticPr, XmlnsAttrs};
use crate::xml::facade::{EditRow, XmlIo};
use crate::xml::sheet_data::SheetData;

#[derive(Debug, Deserialize, Serialize)]
#[serde(rename="worksheet")]
pub(crate) struct WorkSheet {
    #[serde(flatten)]
    xmlns_attrs: XmlnsAttrs,
    #[serde(rename = "dimension", skip_serializing_if = "Option::is_none")]
    dimension: Option<Dimension>,
    #[serde(rename = "sheetViews")]
    sheet_views: SheetViews,
    #[serde(rename = "sheetFormatPr")]
    sheet_format_pr: SheetFormatPr,
    #[serde(rename = "cols", skip_serializing_if = "Option::is_none")]
    cols: Option<Cols>,
    #[serde(rename = "sheetData")]
    pub(crate) sheet_data: SheetData,
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
    #[serde(rename = "Selection", skip_serializing_if = "Option::is_none")]
    selection: Option<Selection>
}

#[derive(Debug, Deserialize, Serialize)]
struct Selection {
    #[serde(rename = "@activeCell")]
    active_cell: String,
    #[serde(rename = "@sqref")]
    sqref: String
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

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct Col {
    #[serde(rename = "@min")]
    min: u32,
    #[serde(rename = "@max")]
    max: u32,
    #[serde(rename = "@customWidth")]
    custom_width: bool,
    #[serde(rename = "@width")]
    width: f64,
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct Cols {
    col: Vec<Col>
}

impl WorkSheet {
    pub(crate) fn from_path<P: AsRef<Path>>(file_path: P, sheet_id: u32) -> WorkSheet {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::SheetFile(sheet_id)).unwrap();
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let work_sheet = de::from_str(&xml).unwrap();
        work_sheet
    }

    pub(crate) fn save<P: AsRef<Path>>(&mut self, file_path: P, sheet_id: u32) {
        self.sheet_data.sort();
        let xml = se::to_string_with_root("worksheet", &self).unwrap();
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::SheetFile(sheet_id)).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}