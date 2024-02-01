use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::common::XmlnsAttrs;
use crate::xml::facade::XmlIo;

#[derive(Debug, Deserialize, Serialize)]
#[serde(rename="workbook")]
pub(crate) struct Workbook {
    #[serde(flatten)]
    xmlns_attrs: XmlnsAttrs,
    #[serde(rename = "fileVersion")]
    file_version: FileVersion,
    #[serde(rename = "workbookPr")]
    workbook_pr: WorkbookPr,
    #[serde(rename = "bookViews")]
    book_views: BookViews,
    #[serde(rename = "sheets")]
    pub(crate) sheets: Sheets,
    #[serde(rename = "calcPr")]
    calc_pr: CalcPr,
    #[serde(rename = "extLst")]
    ext_lst: ExtLst,
}

#[derive(Debug, Deserialize, Serialize)]
struct FileVersion {
    #[serde(rename = "@appName")]
    app_name: String,
    #[serde(rename = "@lastEdited", skip_serializing_if = "Option::is_none")]
    last_edited: Option<u32>,
    #[serde(rename = "@lowestEdited", skip_serializing_if = "Option::is_none")]
    lowest_edited: Option<u32>,
    #[serde(rename = "@rupBuild", skip_serializing_if = "Option::is_none")]
    rup_build: Option<String>,
}

#[derive(Debug, Deserialize, Serialize)]
struct WorkbookPr {
    #[serde(rename = "@filterPrivacy")]
    filter_privacy: u32,
    #[serde(rename = "@defaultThemeVersion")]
    default_theme_version: String,
}

#[derive(Debug, Deserialize, Serialize)]
struct BookViews {
    #[serde(rename = "workbookView")]
    book_views: Vec<WorkbookView>
}

#[derive(Debug, Deserialize, Serialize)]
struct WorkbookView {
    #[serde(rename = "@xWindow")]
    x_window: u32,
    #[serde(rename = "@yWindow")]
    y_window: u32,
    #[serde(rename = "@windowWidth")]
    window_width: u32,
    #[serde(rename = "@windowHeight")]
    window_height: u32,
    #[serde(rename = "@activeTab")]
    active_tab: u32
}

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Sheets {
    #[serde(rename = "sheet")]
    pub(crate) sheets: Vec<Sheet>
}

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Sheet {
    #[serde(rename = "@name")]
    pub(crate) name: String,
    #[serde(rename = "@sheetId")]
    pub(crate) sheet_id: u32,
    #[serde(rename(serialize = "@r:id", deserialize = "@id"))]
    r_id: String,
}

#[derive(Debug, Deserialize, Serialize)]
struct CalcPr {
    #[serde(rename = "@calcId")]
    calc_id: String,
}

#[derive(Debug, Deserialize, Serialize)]
struct ExtLst {
    #[serde(rename = "ext")]
    ext: Vec<Ext>,
}

#[derive(Debug, Deserialize, Serialize)]
struct Ext {
    #[serde(rename = "@uri")]
    uri: String,
    #[serde(rename = "@xmlns:x15", skip_serializing_if = "Option::is_none")]
    xmlns_x15: Option<String>,
    #[serde(rename(serialize = "x15:workbookPr", deserialize = "workbookPr"), skip_serializing_if = "Option::is_none")]
    x15workbook_pr: Option<X15WorkbookPr>,
}

#[derive(Debug, Deserialize, Serialize)]
struct X15WorkbookPr {
    #[serde(rename = "@chartTrackingRefBase", skip_serializing_if = "Option::is_none")]
    chart_tracking_ref_base: Option<u32>
}


impl XmlIo<Workbook> for Workbook {
    fn from_path<P: AsRef<Path>>(file_path: P) -> Workbook {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::WorkbookFile).unwrap();
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let work_book = de::from_str(&xml).unwrap();
        work_book
    }

    fn save<P: AsRef<Path>>(&mut self, file_path: P) {
        let xml = se::to_string_with_root("workbook", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::WorkbookFile).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}