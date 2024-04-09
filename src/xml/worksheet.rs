use std::collections::HashMap;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::api::cell::location::{Location, LocationRange};
use crate::api::relationship::Rel;
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::{Filters, FormatColor};
use crate::api::worksheet::col::ColSet;
use crate::result::{ColResult, WorkSheetResult};
use crate::xml::common::{PhoneticPr, XmlnsAttrs};
use crate::xml::worksheet::auto_filter::AutoFilter;
use crate::xml::worksheet::columns::Cols;
use crate::xml::worksheet::hyperlinks::Hyperlinks;
use crate::xml::worksheet::ignore_errors::IgnoredErrors;
use crate::xml::worksheet::merge_cells::MergeCells;
use crate::xml::worksheet::page_margins::PageMargins;
use crate::xml::worksheet::row_breaks::RowBreaks;
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
mod row_breaks;

#[derive(Debug, Clone, Deserialize, Serialize)]
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
    merge_cells: Option<MergeCells>,
    #[serde(rename = "phoneticPr", skip_serializing_if = "Option::is_none")]
    phonetic_pr: Option<PhoneticPr>,
    #[serde(rename = "hyperlinks", skip_serializing_if = "Option::is_none")]
    hyperlinks: Option<Hyperlinks>,
    #[serde(rename = "autoFilter", skip_serializing_if = "Option::is_none")]
    auto_filter: Option<AutoFilter>,
    #[serde(rename = "printOptions", skip_serializing_if = "Option::is_none")]
    print_options: Option<PrintOptions>,
    #[serde(rename = "pageMargins")]
    page_margins: PageMargins,
    #[serde(rename = "pageSetup", skip_serializing_if = "Option::is_none")]
    page_setup: Option<PageSetup>,
    #[serde(rename = "autoFilter", skip_serializing_if = "Option::is_none")]
    header_footer: Option<HeaderFooter>,
    #[serde(rename = "rowBreaks", skip_serializing_if = "Option::is_none")]
    row_breakers: Option<RowBreaks>,
    #[serde(rename = "ignoredErrors", skip_serializing_if = "Option::is_none")]
    ignored_errors: Option<IgnoredErrors>,
    #[serde(rename = "legacyDrawing", skip_serializing_if = "Option::is_none")]
    legacy_drawing: Option<Drawing>,
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
    pub(crate) fn set_col_by_colset<R: LocationRange>(&mut self, col_range: R, col_set: &ColSet) -> ColResult<()> {
        let (min, max) = col_range.to_col_range();
        let col = self.cols.get_or_new_col(min, max);
        if let Some(width) = col_set.width {
            col.width = Some(width);
            col.custom_width = Some(1);
        }
        if let Some(style) = col_set.style {
            col.style = Some(style);
        }
        if let Some(hidden) = col_set.hidden {
            col.hidden = Some(hidden)
        }
        if let Some(outline_level) = col_set.outline_level {
            col.outline_level = Some(outline_level)
        }
        if let Some(collapsed) = col_set.collapsed {
            self.sheet_format_pr.set_outline_level_col(col.outline_level.unwrap_or(0) as u8);
            col.collapsed = Some(collapsed)
        }
        Ok(())
    }

    pub(crate) fn get_default_style<L: Location>(&self, loc: &L) -> Option<u32> {
        let cell_style = self.sheet_data.get_default_style(loc);
        if let None = cell_style {
            let col_style = self.cols.get_default_style(loc.to_col());
            return col_style;
        }
        cell_style
    }

    pub(crate) fn add_merge_cell(&mut self, first_row: u32, first_col: u32, last_row: u32, last_col: u32) {
        let merge_cells = self.merge_cells.get_or_insert(Default::default());
        merge_cells.add_merge_cell(first_row, first_col, last_row, last_col);
    }

    pub(crate) fn autofit_cols(&mut self) {
        // self.cols.col.iter_mut().for_each(|c| {
        //     c.custom_width = 0;
        //     c.width = None;
        //     c.best_fit = Some(1)
        // })
    }

    pub(crate) fn set_tab_color(&mut self, tab_color: &FormatColor) {
        let sheet_pr = self.sheet_pr.get_or_insert(Default::default());
        sheet_pr.set_tab_color(tab_color);
    }

    pub(crate) fn set_background(&mut self, r_id: u32) {
        self.picture = Some(Picture::from_id(r_id));
    }

    pub(crate) fn insert_image(&mut self, r_id: u32) {
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

    pub(crate) fn get_default_row_height(&self) -> f64 {
        self.sheet_format_pr.get_default_row_height()
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

#[derive(Debug, Clone, Deserialize, Serialize)]
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

#[derive(Debug, Clone, Deserialize, Serialize)]
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
            page_setup: None,
            header_footer: None,
            print_options: None,
            row_breakers: None,
            ignored_errors: None,
            picture: None,
            hyperlinks: None,
            drawing: None,
            auto_filter: None,
            legacy_drawing: None,
        }
    }
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct PageSetup {
    #[serde(rename = "@paperSize", skip_serializing_if = "Option::is_none")]
    paper_size: Option<u8>,
    #[serde(rename = "@scale", skip_serializing_if = "Option::is_none")]
    scale: Option<u32>,
    #[serde(rename = "@orientation", skip_serializing_if = "Option::is_none")]
    orientation: Option<String>,
    #[serde(rename = "@horizontalDpi", skip_serializing_if = "Option::is_none")]
    horizontal_dpi: Option<i32>,
    #[serde(rename = "@verticalDpi", skip_serializing_if = "Option::is_none")]
    vertical_dpi: Option<i32>,
    #[serde(rename(serialize = "@r:id", deserialize = "@id"), skip_serializing_if = "Option::is_none")]
    r_id: Option<Rel>,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct HeaderFooter {
    #[serde(rename = "@alignWithMargins", skip_serializing_if = "Option::is_none")]
    align_with_margins: Option<u8>,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct PrintOptions {
    #[serde(rename = "@horizontalCentered", skip_serializing_if = "Option::is_none")]
    horizontal_centered: Option<u8>,
}

impl WorkSheet {
    pub(crate) fn from_path<P: AsRef<Path>>(file_path: P, target: &str) -> WorkSheetResult<WorkSheet> {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::SheetFile(target.to_string()))?;
        let mut xml = String::new();
        file.read_to_string(&mut xml)?;
        let work_sheet = de::from_str(&xml).unwrap();
        Ok(work_sheet)
    }

    pub(crate) fn save<P: AsRef<Path>>(& self, file_path: P, target: &str) {
        let xml = se::to_string_with_root("worksheet", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::SheetFile(target.to_string())).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
struct Drawing {
    #[serde(rename(serialize = "@r:id", deserialize = "@id"))]
    r_id: Rel,
}