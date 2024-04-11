pub(crate) mod write;
pub(crate) mod row;
pub(crate) mod col;
pub(crate) mod read;
pub(crate) mod column;
mod format;
mod hyperlink;
mod image;

use std::cell::RefCell;
use std::collections::HashMap;
use std::fs::File;
use std::path::Path;
use std::rc::Rc;
use zip::result::ZipResult;
use crate::{Filters, FormatColor, WorkbookResult, xml};
use crate::api::cell::location::{Location, LocationRange};
use crate::api::worksheet;
use crate::api::worksheet::col::Col;
use crate::api::worksheet::image::_Image;
use crate::api::worksheet::read::Read;
use crate::api::worksheet::row::Row;
use crate::api::worksheet::write::Write;
use crate::file::XlsxFileType;
use crate::result::{WorkSheetError, WorkSheetResult};
use crate::xml::drawings::Drawings;
use crate::xml::drawings::vml_drawing::VmlDrawing;
use crate::xml::metadata::Metadata;
use crate::xml::relationships::Relationships;
use crate::xml::shared_string::SharedString;
use crate::xml::worksheet::WorkSheet as XmlWorkSheet;
use crate::xml::workbook::Workbook;
use crate::xml::style::StyleSheet;

#[derive(Debug)]
pub struct WorkSheet {
    pub(crate) id: u32,
    pub(crate) name: String,
    pub(crate) target: String,
    workbook: Rc<RefCell<Workbook>>,
    workbook_rel: Rc<RefCell<Relationships>>,
    // workbook_api: Weak<& mut ApiWorkbook>,
    worksheet: XmlWorkSheet,
    worksheet_rel: Relationships,
    style_sheet: Rc<RefCell<StyleSheet>>,
    content_types: Rc<RefCell<xml::content_types::ContentTypes>>,
    medias: Rc<RefCell<xml::medias::Medias>>,
    vml_drawing: Option<VmlDrawing>,
    drawings: Option<Drawings>,
    drawings_rel: Option<Relationships>,
    metadata: Rc<RefCell<Metadata>>,
    shared_string: Rc<SharedString>,
}

impl Write for WorkSheet {}
impl Row for WorkSheet {
    // fn hide_row(&mut self, row: u32) -> WorkSheetResult<()> {
    //     self.worksheet.sheet_data.hide_row(row);
    //     Ok(())
    // }
}
impl Col for WorkSheet {}
impl Read for WorkSheet {}

impl WorkSheet {
    pub(crate) fn save_as<P: AsRef<Path>>(&self, file_path: P) -> WorkSheetResult<()> {
        self.worksheet.save(&file_path, &self.target);
        self.worksheet_rel.save(&file_path, XlsxFileType::WorksheetRels(self.id));
        if let Some(_) = self.worksheet_rel.get_drawings_rid() {
            self.drawings.as_ref().unwrap().save(&file_path, self.id);
            self.drawings_rel.as_ref().unwrap().save(&file_path, XlsxFileType::DrawingRels(self.id));
        }
        if let Some(id) = self.worksheet_rel.get_vml_drawing_rid() {
            self.vml_drawing.as_ref().unwrap().save(&file_path, id);
        }
        Ok(())
    }
}

impl WorkSheet {
    pub fn autofilter<L: LocationRange>(&mut self, loc_range: L) {
        self.worksheet.autofilter(loc_range);
    }

    pub fn filter_column<L: Location>(&mut self, col: L, filters: &Filters) {
        self.worksheet.filter_column(col, filters);
    }
}

impl WorkSheet {
    pub fn max_column(&self) -> u32 {
        let worksheet = &self.worksheet;
        let sheet_data = &worksheet.sheet_data;
        sheet_data.max_col()
    }

    pub fn max_row(&self) -> u32 {
        let worksheet = &self.worksheet;
        let sheet_data = &worksheet.sheet_data;
        sheet_data.max_row()
    }

    // fn autofit(&mut self) {
    //     todo!();
    //     let worksheets = &mut self.worksheets.borrow_mut();
    //     let worksheet = worksheets.get_mut(&self.id).unwrap();
    //     worksheet.autofit_cols();
    // }

    pub fn get_name(&self) -> &str {
        &self.name
    }

    pub fn set_name(&mut self, name: &str) -> WorkSheetResult<()> {
        let workbook = &mut self.workbook.borrow_mut();
        let name_existed = workbook.sheets.sheets
            .iter()
            .filter(|sheet| sheet.name == name)
            .count() > 0;
        if name_existed {
            Err(WorkSheetError::DuplicatedSheets)
        } else {
            workbook.sheets.sheets
                .iter_mut()
                .filter(|sheet| sheet.name == self.name)
                .for_each(|sheet| sheet.name = name.to_string());
            self.name = name.to_string();
            Ok(())
        }
    }

    pub fn activate(&mut self) {
        let workbook = &mut self.workbook.borrow_mut();
        let book_views = &mut workbook.book_views;
        book_views.set_active_tab(self.id - 1)
    }

    pub fn select(&mut self) {
        let worksheet = &mut self.worksheet;
        let sheet_views = &mut worksheet.sheet_views;
        sheet_views.set_tab_selected(1);
    }

    pub fn deselect(&mut self) {
        let worksheet = &mut self.worksheet;
        let sheet_views = &mut worksheet.sheet_views;
        sheet_views.set_tab_selected(0);
    }

    pub fn right_to_left(&mut self) {
        self.worksheet.sheet_views.set_right_to_left(1);
    }

    pub fn set_top_left_cell<L: Location>(&mut self, loc: L) {
        let worksheet = &mut self.worksheet;
        let sheet_views = &mut worksheet.sheet_views;
        sheet_views.set_top_left_cell(&loc.to_ref());
    }

    pub fn set_zoom(&mut self, zoom_scale: u16) {
        let worksheet = &mut self.worksheet;
        let sheet_views = &mut worksheet.sheet_views;
        sheet_views.set_zoom_scale(zoom_scale);
    }

    pub fn set_selection<L: LocationRange>(&mut self, loc_range: L) -> WorkSheetResult<()> {
        let worksheet = &mut self.worksheet;
        let sheet_views = &mut worksheet.sheet_views;
        sheet_views.set_selection(&loc_range);
        Ok(())
    }

    pub fn freeze_panes<L: Location>(&mut self, loc: L) -> WorkSheetResult<()> {
        let worksheet = &mut self.worksheet;
        let sheet_views = &mut worksheet.sheet_views;
        sheet_views.set_freeze_panes(loc);
        Ok(())
    }

    pub fn split_panes(&mut self, width: f64, height: f64) -> WorkSheetResult<()> {
        let width = (420.0 + width * 120.0) as u32;
        let height = (280.0 + height * 20.0) as u32;
        let worksheet = &mut self.worksheet;
        let sheet_views = &mut worksheet.sheet_views;
        sheet_views.split_panes(width, height);
        Ok(())
    }

    pub fn set_default_row(&mut self, height: f64) {
        let worksheet = &mut self.worksheet;
        worksheet.set_default_row_height(height);
    }

    pub fn get_default_row(&self) -> f64 {
        let worksheet = &self.worksheet;
        worksheet.get_default_row_height()
    }

    pub fn hide_unused_rows(&mut self, hide: bool) {
        let worksheet = &mut self.worksheet;
        worksheet.hide_unused_rows(hide);
    }

    pub fn outline_settings(&mut self, visible: bool, symbols_below: bool, symbols_right: bool, auto_style: bool) {
        let worksheet = &mut self.worksheet;
        worksheet.outline_settings(visible, symbols_below, symbols_right, auto_style)
    }

    pub fn ignore_errors<L: Location>(&mut self, error_map: HashMap<&str, L>) {
        let error_map = error_map.iter().map(|(&e, l)| (e, l.to_ref())).collect::<HashMap<&str, String>>();
        let worksheet = &mut self.worksheet;
        worksheet.ignore_errors(error_map);
    }

    pub fn hide(&mut self) {
        let mut workbook = self.workbook.borrow_mut();
        let sheet = &mut workbook
            .sheets.sheets
            .iter_mut()
            .find(|s| { s.sheet_id == self.id })
            .unwrap();
        sheet.state = Some(String::from("hidden"));
    }

    fn set_first_sheet(&mut self) {
        self.change_id(1).unwrap();
    }

    fn change_id(&mut self, id: u32) -> WorkbookResult<()> {
        // swap sheet
        // let sheet_target = self.worksheets.borrow_mut().remove(&id).ok_or(SheetError::FileNotFound)?;
        // let sheet = self.worksheets.borrow_mut().remove(&self.id).ok_or(SheetError::FileNotFound)?;
        // self.worksheets.borrow_mut().insert(id, sheet);
        // self.worksheets.borrow_mut().insert(self.id, sheet_target);
        // // swap sheet rel
        // let sheet_rel_target = self.worksheets_rel.borrow_mut().remove(&id);
        // let sheet_rel = self.worksheets_rel.borrow_mut().remove(&self.id);
        // if let Some(sheet_rel_target) = sheet_rel_target {
        //     self.worksheets_rel.borrow_mut().insert(self.id, sheet_rel_target);
        // }
        // if let Some(sheet_rel) = sheet_rel {
        //     self.worksheets_rel.borrow_mut().insert(id, sheet_rel);
        // }
        // // swap workbook sheet
        // let loc = (self.id - 1) as usize;
        // let workbook = &mut self.workbook.borrow_mut();
        // let sheets = &mut workbook.sheets.sheets;
        // sheets[loc].change_id(id);
        // sheets[0].change_id(self.id);
        // sheets.swap(0, loc);
        // // change self id
        Ok(())
    }

    pub fn set_tab_color(&mut self, tab_color: &FormatColor) {
        self.worksheet.set_tab_color(tab_color);
    }

    pub fn set_background<P: AsRef<Path>>(&mut self, filename: P) -> WorkSheetResult<()> {
        let r_id = self.add_background(&filename)?;
        self.worksheet.set_background(r_id);
        Ok(())
    }

    pub fn insert_image<L: LocationRange, P: AsRef<Path>>(&mut self, loc_range: L, filename: &P) -> WorkSheetResult<()> {
        let (from_row, from_col, to_row, to_col) = loc_range.to_range();
        let r_id = self.add_drawing((from_row, from_col, to_row, to_col), filename)?;
        self.worksheet.insert_image(r_id);
        Ok(())
    }

    pub fn id(&self) -> u32 {
        self.id
    }
}

impl WorkSheet {
    // pub(crate) fn add_drawings(&mut self, tmp_path: &str) {
    // let worksheet_rel = &mut self.worksheet_rel;
    // let drawings_id = worksheet_rel.list_drawings();
    // drawings_id.iter().for_each(|&id| {
    //     self.drawings.insert(id, Drawings::from_path(tmp_path, id).unwrap());
    // });
    // }

    pub(crate) fn from_worksheet<P: AsRef<Path>>(
        sheet_id: u32,
        name: &str,
        target: &str,
        tmp_path: P,
        workbook: Rc<RefCell<Workbook>>,
        workbook_rel: Rc<RefCell<Relationships>>,
        style_sheet: Rc<RefCell<StyleSheet>>,
        content_types: Rc<RefCell<xml::content_types::ContentTypes>>,
        medias: Rc<RefCell<xml::medias::Medias>>,
        metadata: Rc<RefCell<Metadata>>,
        worksheet: &WorkSheet,
        shared_string: Rc<SharedString>
    ) -> WorkSheet {
        let mut new_worksheet = Self::from_xml(
            sheet_id,
            name,
            target,
            tmp_path,
            workbook,
            workbook_rel,
            style_sheet,
            content_types,
            medias,
            metadata,
            shared_string,
        );
        new_worksheet.worksheet = worksheet.worksheet.clone();
        new_worksheet.worksheet_rel = worksheet.worksheet_rel.clone();
        new_worksheet.drawings = worksheet.drawings.clone();
        new_worksheet.drawings_rel = worksheet.drawings_rel.clone();
        new_worksheet
    }

    pub(crate) fn from_xml<P: AsRef<Path>>(
        sheet_id: u32,
        name: &str,
        target: &str,
        file_path: P,
        workbook: Rc<RefCell<Workbook>>,
        workbook_rel: Rc<RefCell<Relationships>>,
        style_sheet: Rc<RefCell<StyleSheet>>,
        content_types: Rc<RefCell<xml::content_types::ContentTypes>>,
        medias: Rc<RefCell<xml::medias::Medias>>,
        metadata: Rc<RefCell<Metadata>>,
        shared_string: Rc<SharedString>,
    ) -> WorkSheet {
        let file = File::open(&file_path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        // Read worksheet from zip dir
        let mut worksheet = XmlWorkSheet::default();
        let mut worksheet_rel = Relationships::default();
        {
            let worksheet_file = archive.by_name(&format!("xl/{target}"));
            if let Ok(mut worksheet_file) = worksheet_file {
                worksheet = XmlWorkSheet::from_zip_file(&mut worksheet_file);
                // Prevent incorrect results from being filled into cells
                worksheet.sheet_data.clean_formula_value();
            };
        }
        {
            if let Ok(mut worksheet_rel_file) = archive.by_name(&format!("xl/worksheets/_rels/sheet{sheet_id}.xml.rels")) {
                worksheet_rel = Relationships::from_zip_file(&mut worksheet_rel_file);
            }
        }
        // load drawings
        let drawings_id = worksheet_rel.get_drawings_rid();
        let (mut drawings, mut drawings_rel) = (None, None);
        if let Some(drawings_id) = drawings_id {
            {
                let mut drawings_file = archive.by_name(&format!("xl/drawings/drawing{drawings_id}.xml")).unwrap();
                drawings = Some(Drawings::from_zip_file(&mut drawings_file));
            }
            {
                let mut drawings_rel_file = archive.by_name(&format!("xl/drawings/_rels/drawing{drawings_id}.xml.rels")).unwrap();
                drawings_rel = Some(Relationships::from_zip_file(&mut drawings_rel_file));
            }
        };
        let vml_drawing = match worksheet_rel.get_vml_drawing_rid() {
            Some(vml_drawing_id) => VmlDrawing::from_path(&file_path, vml_drawing_id).ok(),
            None => None
        };
        WorkSheet {
            id: sheet_id,
            name: String::from(name),
            target: format!("{target}"),
            workbook,
            workbook_rel,
            worksheet,
            worksheet_rel,
            style_sheet,
            content_types,
            medias,
            vml_drawing,
            drawings,
            drawings_rel,
            metadata,
            shared_string,
        }
    }
}