use std::collections::HashMap;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::api::cell::location::{Location, LocationRange};
use crate::api::relationship::Rel;
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::{Filters, FormatColor};
use crate::result::{ColResult, WorkSheetResult};
use crate::xml::common::{PhoneticPr, XmlnsAttrs};
use crate::xml::worksheet::auto_filter::AutoFilter;
use crate::xml::worksheet::columns::Cols;
use crate::xml::worksheet::hyperlinks::Hyperlinks;
use crate::xml::worksheet::ignore_errors::IgnoredErrors;
use crate::xml::worksheet::merge_cells::MergeCells;
use crate::xml::worksheet::page_margins::PageMargins;
use crate::xml::worksheet::sheet_format::SheetFormatPr;
use self::sheet_views::SheetViews;
use self::sheet_data::SheetData;
use self::sheet_pr::SheetPr;

mod sheet_data;
mod sheet_pr;
mod sheet_format;
mod sheet_views;
mod merge_cells;
mod columns;
mod ignore_errors;
mod hyperlinks;
mod page_margins;
mod auto_filter;

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
    cols: Cols,
    #[serde(rename = "sheetData", default)]
    pub(crate) sheet_data: SheetData,
    #[serde(rename = "mergeCells", skip_serializing_if = "Option::is_none")]
    merge_cells: Option<MergeCells>,
    #[serde(rename = "phoneticPr", skip_serializing_if = "Option::is_none")]
    phonetic_pr: Option<PhoneticPr>,
    #[serde(rename = "hyperlinks", skip_serializing_if = "Option::is_none")]
    hyperlinks: Option<Hyperlinks>,
    #[serde(rename = "autoFilter", skip_serializing_if = "Option::is_none")]
    auto_filter: Option<AutoFilter>,
    #[serde(rename = "pageMargins")]
    page_margins: PageMargins,
    #[serde(rename = "ignoredErrors", skip_serializing_if = "Option::is_none")]
    ignored_errors: Option<IgnoredErrors>,
    #[serde(rename = "drawing", skip_serializing_if = "Option::is_none")]
    drawing: Option<Drawing>,
    #[serde(rename = "picture", skip_serializing_if = "Option::is_none")]
    picture: Option<Picture>,
}

impl WorkSheet {
    pub(crate) fn autofilter<L: LocationRange>(&mut self, loc_range: L) {
        let auto_filter = self.auto_filter.get_or_insert(AutoFilter::default());
        auto_filter.sqref = loc_range.to_range_ref();
    }

    pub(crate) fn filter_column<L: Location>(&mut self, col: L, filters: &Filters) {
        let auto_filter = self.auto_filter.get_or_insert(AutoFilter::default());
        auto_filter.add_filters(col.to_col(), filters);
    }
}

impl WorkSheet {
    pub(crate) fn create_col(&mut self, min: u32, max: u32, width: Option<f64>, style: Option<u32>, best_fit: Option<u8>) -> ColResult<()> {
        self.cols.update_col(min, max, width, style, None, best_fit)
    }

    pub(crate) fn hide_col(&mut self, min: u32, max: u32, hidden: Option<u8>) -> ColResult<()> {
        self.cols.update_col(min, max, None, None, hidden, None)
    }

    pub(crate) fn add_merge_cell(&mut self, first_row: u32, first_col: u32, last_row: u32, last_col: u32) {
        let mut merge_cells = self.merge_cells.take().unwrap_or_default();
        merge_cells.add_merge_cell(first_row, first_col, last_row, last_col);
        self.merge_cells = Some(merge_cells);
    }

    pub(crate) fn autofit_cols(&mut self) {
        // self.cols.col.iter_mut().for_each(|c| {
        //     c.custom_width = 0;
        //     c.width = None;
        //     c.best_fit = Some(1)
        // })
    }

    pub(crate) fn set_tab_color(&mut self, tab_color: &FormatColor) {
        let mut sheet_pr = self.sheet_pr.take().unwrap_or_default();
        sheet_pr.set_tab_color(tab_color);
        self.sheet_pr = Some(sheet_pr);
    }

    pub(crate) fn set_background(&mut self, r_id: u32) {
        self.picture = Some(Picture::from_id(r_id));
    }

    pub(crate) fn insert_image<L: Location>(&mut self, loc: L, r_id: u32) {
        let drawing = self.drawing.get_or_insert(Default::default());
        drawing.r_id = Rel::from_id(r_id);
    }

    pub(crate) fn add_hyperlink<L: Location>(&mut self, loc: &L, r_id: u32) {
        let hyperlinks = self.hyperlinks.get_or_insert(Default::default());
        hyperlinks.add_hyperlink(loc, r_id);
    }

    pub(crate) fn set_default_row_height(&mut self, height: f64) {
        self.sheet_format_pr.set_default_row_height(height);
    }

    pub(crate) fn hide_unused_rows(&mut self, hide: bool) {
        self.sheet_format_pr.hide_unused_rows(hide);
    }

    pub(crate) fn outline_settings(&mut self, visible: bool, symbols_below: bool, symbols_right: bool, auto_style: bool) {
        let sheet_pr = self.sheet_pr.get_or_insert(Default::default());
        sheet_pr.set_outline_pr(visible, symbols_below, symbols_right, auto_style);
    }

    pub(crate) fn ignore_errors(&mut self, error_map: HashMap<&str, String>) {
        let ignore_errors = IgnoredErrors::from_map(error_map);
        self.ignored_errors = Some(ignore_errors);
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

impl Default for WorkSheet {
    fn default() -> WorkSheet {
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
            drawing: None,
            auto_filter: None,
        }
    }
}

impl WorkSheet {
    pub(crate) fn from_path<P: AsRef<Path>>(file_path: P, sheet_id: u32) -> WorkSheetResult<WorkSheet> {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::SheetFile(sheet_id))?;
        let mut xml = String::new();
        file.read_to_string(&mut xml)?;
        let work_sheet = de::from_str(&xml).unwrap();
        Ok(work_sheet)
    }

    pub(crate) fn save<P: AsRef<Path>>(&mut self, file_path: P, sheet_id: u32) {
        let xml = se::to_string_with_root("worksheet", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::SheetFile(sheet_id)).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}

#[derive(Debug, Deserialize, Serialize, Default)]
struct Drawing {
    #[serde(rename(serialize = "@r:id", deserialize = "@id"))]
    r_id: Rel,
}