use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::FormatColor;
use crate::result::ColResult;
use crate::utils::col_helper::to_ref;
use crate::xml::common::{FromFormat, PhoneticPr, XmlnsAttrs};
use crate::xml::merge_cells::MergeCells;
use crate::xml::sheet_data::SheetData;
use crate::xml::style::color::Color;

#[derive(Debug, Deserialize, Serialize)]
#[serde(rename="worksheet")]
pub(crate) struct WorkSheet {
    #[serde(flatten)]
    xmlns_attrs: XmlnsAttrs,
    #[serde(rename = "sheetPr", skip_serializing_if = "Option::is_none")]
    sheet_pr: Option<SheetPr>,
    #[serde(rename = "dimension", skip_serializing_if = "Option::is_none")]
    dimension: Option<Dimension>,
    #[serde(rename = "sheetViews")]
    pub(crate) sheet_views: SheetViews,
    #[serde(rename = "sheetFormatPr")]
    sheet_format_pr: SheetFormatPr,
    #[serde(rename = "cols", default, skip_serializing_if = "Cols::is_empty")]
    pub(crate) cols: Cols,
    #[serde(rename = "sheetData", default)]
    pub(crate) sheet_data: SheetData,
    #[serde(rename = "mergeCells", skip_serializing_if = "Option::is_none")]
    pub(crate) merge_cells: Option<MergeCells>,
    #[serde(rename = "phoneticPr", skip_serializing_if = "Option::is_none")]
    phonetic_pr: Option<PhoneticPr>,
    #[serde(rename = "hyperlinks", skip_serializing_if = "Option::is_none")]
    hyperlinks: Option<Hyperlinks>,
    #[serde(rename = "pageMargins")]
    page_margins: PageMargins,
    #[serde(rename = "picture", skip_serializing_if = "Option::is_none")]
    picture: Option<Picture>,
}

impl WorkSheet {
    pub(crate) fn create_col(&mut self, min: u32, max: u32, width: Option<f64>, style: Option<u32>, best_fit: u8) -> ColResult<&mut Col> {
        let mut col = Col::new(min, max, 1, width, style, best_fit);
        if let None = width {
            col.custom_width = 0;
        }
        self.cols.col.push(col);
        Ok(self.cols.col.last_mut().unwrap())
    }

    pub(crate) fn add_merge_cell(&mut self, first_row: u32, first_col: u32, last_row: u32, last_col: u32) {
        let mut merge_cells = self.merge_cells.take().unwrap_or_default();
        merge_cells.add_merge_cell(first_row, first_col, last_row, last_col);
        self.merge_cells = Some(merge_cells);
    }

    pub(crate) fn autofit_cols(&mut self) {
        self.cols.col.iter_mut().for_each(|c| {
            c.custom_width = 0;
            c.width = None;
            c.best_fit = 1
        })
    }

    pub(crate) fn set_tab_color(&mut self, tab_color: &FormatColor) {
        let sheet_pr = match self.sheet_pr.take() {
            Some(mut sheet_pr) => {
                sheet_pr.tab_color = Color::from_format(tab_color);
                sheet_pr
            },
            None => {
                SheetPr::new(tab_color)
            }
        };
        self.sheet_pr = Some(sheet_pr);
    }

    pub(crate) fn set_background(&mut self, r_id: u32) {
        if let None = self.picture {
            self.picture = Some(Picture::from_id(r_id))
        }
    }

    pub(crate) fn add_hyperlink(&mut self, row: u32, col: u32, r_id: u32) {
        let hyperlink = Hyperlink::new(&to_ref(row, col), r_id);
        if let None = self.hyperlinks {
            let mut hyperlinks = Hyperlinks::default();
            hyperlinks.add_hyperlink(hyperlink);
            self.hyperlinks = Some(hyperlinks);
        } else {
            self.hyperlinks.as_mut().unwrap().add_hyperlink(hyperlink);
        };
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct SheetPr {
    #[serde(rename = "tabColor")]
    tab_color: Color
}

impl SheetPr {
    fn new(color: &FormatColor) -> SheetPr {
        SheetPr {
            tab_color: Color::from_format(color),
        }
    }
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
pub(crate) struct SheetView {
    #[serde(rename = "@tabSelected", skip_serializing_if = "Option::is_none")]
    pub(crate) tab_selected: Option<u32>,
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
pub(crate) struct SheetViews {
    #[serde(rename = "sheetView")]
    pub(crate) sheet_view: Vec<SheetView>
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

impl Default for PageMargins {
    fn default() -> PageMargins {
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

#[derive(Debug, Deserialize, Serialize)]
struct Picture {
    #[serde(rename(serialize = "@r:id", deserialize = "@id"))]
    r_id: String,
}

impl Picture {
    fn from_id(r_id: u32) -> Picture {
        Picture {
            r_id: format!("rId{r_id}"),
        }
    }
}

#[derive(Debug, Deserialize, Serialize, PartialEq, Default)]
pub struct Col {
    #[serde(rename = "@min")]
    min: u32,
    #[serde(rename = "@max")]
    max: u32,
    #[serde(rename = "@width", skip_serializing_if = "Option::is_none")]
    width: Option<f64>,
    #[serde(rename = "@style", skip_serializing_if = "Option::is_none")]
    style: Option<u32>,
    #[serde(rename = "@bestFit", default)]
    best_fit: u8,
    #[serde(rename = "@customWidth")]
    custom_width: u8,
}

impl Col {
    fn new(min: u32, max: u32, custom_width: u8, width: Option<f64>, style: Option<u32>, best_fit: u8) -> Col {
        Col {
            min,
            max,
            custom_width,
            width,
            style,
            best_fit,
        }
    }
}

#[derive(Debug, Deserialize, Serialize, PartialEq, Default)]
pub struct Cols {
    col: Vec<Col>
}

impl Cols {
    fn add_col(&mut self, col: Col) {
        self.col.push(col)
    }
    fn is_empty(&self) -> bool {
        self.col.is_empty()
    }
}


#[derive(Debug, Deserialize, Serialize, Default)]
struct Hyperlinks {
    #[serde(rename = "hyperlink")]
    hyperlink: Vec<Hyperlink>
}

impl Hyperlinks {
    fn add_hyperlink(&mut self, hyperlink: Hyperlink) {
        self.hyperlink.push(hyperlink)
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct Hyperlink {
    #[serde(rename = "@ref")]
    hyperlink_ref: String,
    #[serde(rename(serialize = "@r:id", deserialize = "@id"))]
    r_id: String,
    // #[serde(rename(serialize = "@xr:uid", deserialize = "@uid"))]
    // uid: String,
}

impl Hyperlink {
    fn new(hyperlink_ref: &str, r_id: u32) -> Self {
        Self {
            hyperlink_ref: String::from(hyperlink_ref),
            r_id: format!("rId{r_id}"),
        }
    }
}

impl WorkSheet {
    pub(crate) fn new() -> WorkSheet {
        WorkSheet {
            xmlns_attrs: XmlnsAttrs::worksheet_default(),
            sheet_pr: None,
            dimension: Some(Dimension::default()),
            sheet_views: SheetViews::default(),
            sheet_format_pr: SheetFormatPr::default(),
            cols: Cols::default(),
            sheet_data: SheetData::default(),
            merge_cells: None,
            phonetic_pr: None,
            page_margins: PageMargins::default(),
            picture: None,
            hyperlinks: None,
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
        self.sheet_data.sort_row();
        let xml = se::to_string_with_root("worksheet", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::SheetFile(sheet_id)).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}