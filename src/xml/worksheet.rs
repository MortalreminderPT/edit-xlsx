use std::collections::HashMap;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::FormatColor;
use crate::result::ColResult;
use crate::utils::col_helper::to_ref;
use crate::xml::common::{FromFormat, PhoneticPr, XmlnsAttrs};
use crate::xml::worksheet::mergecells::MergeCells;
use self::sheetviews::SheetViews;
use self::sheetdata::SheetData;
use self::sheetpr::SheetPr;

mod sheetdata;
mod sheetpr;
mod sheetformat;
mod sheetviews;
mod mergecells;

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
    pub(crate) sheet_format_pr: SheetFormatPr,
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
    #[serde(rename = "ignoredErrors", skip_serializing_if = "Option::is_none")]
    ignored_errors: Option<IgnoredErrors>,
    #[serde(rename = "picture", skip_serializing_if = "Option::is_none")]
    picture: Option<Picture>,
}

impl WorkSheet {
    pub(crate) fn create_col(&mut self, min: u32, max: u32, width: Option<f64>, style: Option<u32>, best_fit: Option<u8>) -> ColResult<&mut Col> {
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
            c.best_fit = Some(1)
        })
    }

    pub(crate) fn set_tab_color(&mut self, tab_color: &FormatColor) {
        let mut sheet_pr = self.sheet_pr.take().unwrap_or_default();
        sheet_pr.set_tab_color(tab_color);
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

    pub(crate) fn set_default_row_height(&mut self, height: f64) {
        self.sheet_format_pr.custom_height = Some(1);
        self.sheet_format_pr.default_row_height = height;
    }

    pub(crate) fn hide_unused_rows(&mut self, hide: bool) {
        self.sheet_format_pr.zero_height = Some(hide as u8);
    }

    pub(crate) fn outline_settings(&mut self, visible: bool, symbols_below: bool, symbols_right: bool, auto_style: bool) {
        let mut sheet_pr = self.sheet_pr.take().unwrap_or_default();
        sheet_pr.set_outline_pr(visible, symbols_below, symbols_right, auto_style);
        self.sheet_pr = Some(sheet_pr);
    }

    pub(crate) fn ignore_errors(&mut self, error_map: HashMap<&str, String>) {
        let ignore_errors = IgnoredErrors::from_map(error_map);
        self.ignored_errors = Some(ignore_errors);
    }
}

#[derive(Debug, Deserialize, Serialize, Default)]
struct IgnoredErrors {
    #[serde(rename = "ignoredError")]
    ignored_error: Vec<IgnoredError>,
}

impl IgnoredErrors {
    fn from_map(error_map: HashMap<&str, String>) -> IgnoredErrors {
        let mut ignored_errors = IgnoredErrors::default();
        error_map.iter().for_each(|(error_type, sqref)|{
            ignored_errors.ignored_error.push(IgnoredError::from_sqref(error_type, sqref));
        });
        ignored_errors
    }
}

#[derive(Debug, Deserialize, Serialize, Default)]
struct IgnoredError {
    #[serde(rename = "@sqref")]
    sqref: String,
    #[serde(rename = "@numberStoredAsText")]
    number_stored_as_text: Option<u8>,
    #[serde(rename = "@evalError")]
    eval_error: Option<u8>,
    #[serde(rename = "@formulaDiffers")]
    formula_differs: Option<u8>,
    #[serde(rename = "@formulaRange")]
    formula_range: Option<u8>,
    #[serde(rename = "@formulaUnlocked")]
    formula_unlocked: Option<u8>,
    #[serde(rename = "@emptyCellReference")]
    empty_cell_reference: Option<u8>,
    #[serde(rename = "@listDataValidation")]
    list_data_validation: Option<u8>,
    #[serde(rename = "@calculatedColumn")]
    calculated_column: Option<u8>,
    #[serde(rename = "@twoDigitTextYear")]
    two_digit_text_year: Option<u8>,
}

impl IgnoredError {
    fn from_sqref(error_type: &str, sqref: &str) -> IgnoredError {
        let mut ignored_error = IgnoredError::default();
        match error_type {
            "number_stored_as_text" => {
                ignored_error.number_stored_as_text = Some(1);
                ignored_error.sqref = String::from(sqref);
            },
            "eval_error" => {
                ignored_error.eval_error = Some(1);
                ignored_error.sqref = String::from(sqref);
            },
            "formula_differs" => {
                ignored_error.formula_differs = Some(1);
                ignored_error.sqref = String::from(sqref);
            },
            "formula_range" => {
                ignored_error.formula_range = Some(1);
                ignored_error.sqref = String::from(sqref);
            },
            "formula_unlocked" => {
                ignored_error.formula_unlocked = Some(1);
                ignored_error.sqref = String::from(sqref);
            },
            "empty_cell_reference" => {
                ignored_error.empty_cell_reference = Some(1);
                ignored_error.sqref = String::from(sqref);
            },
            "list_data_validation" => {
                ignored_error.list_data_validation = Some(1);
                ignored_error.sqref = String::from(sqref);
            },
            "calculated_column" => {
                ignored_error.calculated_column = Some(1);
                ignored_error.sqref = String::from(sqref);
            },
            "two_digit_text_year" => {
                ignored_error.two_digit_text_year = Some(1);
                ignored_error.sqref = String::from(sqref);
            },
            &_ => {}
        }
        ignored_error
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
pub(crate) struct SheetFormatPr {
    #[serde(rename = "@defaultRowHeight")]
    default_row_height: f64,
    #[serde(rename = "@customHeight", skip_serializing_if = "Option::is_none")]
    custom_height: Option<u8>,
    #[serde(rename = "@zeroHeight", skip_serializing_if = "Option::is_none")]
    zero_height: Option<u8>,
}

impl Default for SheetFormatPr {
    fn default() -> SheetFormatPr {
        SheetFormatPr {
            default_row_height: 15.0,
            custom_height: None,
            zero_height: None,
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
    #[serde(rename = "@bestFit", skip_serializing_if = "Option::is_none")]
    best_fit: Option<u8>,
    #[serde(rename = "@customWidth")]
    custom_width: u8,
}

impl Col {
    fn new(min: u32, max: u32, custom_width: u8, width: Option<f64>, style: Option<u32>, best_fit: Option<u8>) -> Col {
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
pub(crate) struct Cols {
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
            ignored_errors: None,
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
