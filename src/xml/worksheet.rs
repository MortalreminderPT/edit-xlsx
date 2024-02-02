use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::common::{PhoneticPr, XmlnsAttrs};
use crate::xml::manage::{EditRow, XmlIo};
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
    #[serde(rename = "sheetData", default = "SheetData::default")]
    pub(crate) sheet_data: SheetData,
    #[serde(rename = "phoneticPr", skip_serializing_if = "Option::is_none")]
    phonetic_pr: Option<PhoneticPr>,
    #[serde(rename = "pageMargins")]
    page_margins: PageMargins,
}

impl WorkSheet {
    // pub(crate) fn borrow_sheet_data(&mut self) -> Option<&mut SheetData> {
    //     match &mut self.sheet_data {
    //         Some(sheet_data) => Some(sheet_data),
    //         None => None
    //     }
    // }
    // pub fn create_sheet_data(&mut self) -> &mut SheetData {
    //     self.sheet_data = Some(SheetData::new());
    //     self.borrow_sheet_data().unwrap()
    // }
    //
    // pub(crate) fn borrow_or_create_sheet_data(&mut self) -> &mut SheetData {
    //     self.sheet_data = match self.sheet_data.take() {
    //         Some(sheet_data) => Some(sheet_data),
    //         None => Some(SheetData::new())
    //     };
    //     self.borrow_sheet_data().unwrap()
    // }
}

#[derive(Debug, Deserialize, Serialize)]
struct Dimension {
    #[serde(rename="@ref")]
    refer: String,
}

impl Dimension {
    pub(crate) fn default() -> Dimension {
        Dimension {
            refer: "A1".to_string(),
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct SheetView {
    #[serde(rename = "@tabSelected", skip_serializing_if = "Option::is_none")]
    tab_selected: Option<u32>,
    #[serde(rename = "@workbookViewId")]
    workbook_view_id: u32,
    #[serde(rename = "selection", skip_serializing_if = "Option::is_none")]
    selection: Option<Selection>
}

impl SheetView {
    pub(crate) fn default() -> SheetView {
        SheetView {
            tab_selected: Some(0),
            workbook_view_id: 0,
            selection: None,
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct Selection {
    #[serde(rename = "@activeCell", skip_serializing_if = "Option::is_none")]
    active_cell: Option<String>,
    #[serde(rename = "@sqref", skip_serializing_if = "Option::is_none")]
    sqref: Option<String>
}

#[derive(Debug, Deserialize, Serialize)]
struct SheetViews {
    #[serde(rename = "sheetView")]
    sheet_view: Vec<SheetView>
}

impl SheetViews {
    pub(crate) fn default() -> SheetViews {
        SheetViews {
            sheet_view: vec![SheetView::default()],
        }
    }
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

impl SheetFormatPr {
    pub(crate) fn default() -> SheetFormatPr {
        SheetFormatPr {
            default_row_height: 15.0,
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct PageMargins {
    #[serde(rename = "@left")]
    left: f64,
    #[serde(rename = "@right")]
    right: f64,
    #[serde(rename = "@top")]
    top: f64,
    #[serde(rename = "@bottom")]
    bottom: f64,
    #[serde(rename = "@header")]
    header: f64,
    #[serde(rename = "@footer")]
    footer: f64,
}

impl PageMargins {
    pub(crate) fn default() -> PageMargins {
        PageMargins {
            left: 0.7,
            right: 0.7,
            top: 0.75,
            bottom: 0.75,
            header: 0.3,
            footer: 0.3,
        }
    }
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
    pub(crate) fn new() -> WorkSheet {
        WorkSheet {
            xmlns_attrs: XmlnsAttrs::default(),
            dimension: Some(Dimension::default()),
            sheet_views: SheetViews::default(),
            sheet_format_pr: SheetFormatPr::default(),
            cols: None,
            sheet_data: SheetData::default(),
            phonetic_pr: None,
            page_margins: PageMargins::default(),
        }
    }
    
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
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::SheetFile(sheet_id)).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}