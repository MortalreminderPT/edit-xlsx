pub(crate) mod write;
pub(crate) mod row;
pub(crate) mod col;
mod format;
mod hyperlink;
mod image;

use std::cell::RefCell;
use std::collections::HashMap;
use std::path::Path;
use std::rc::Rc;
use crate::{FormatColor, WorkbookResult, xml};
use crate::api::cell::location::{Location, LocationRange};
use crate::api::worksheet::col::Col;
use crate::api::worksheet::image::_Image;
use crate::api::worksheet::row::Row;
use crate::api::worksheet::write::Write;
use crate::file::XlsxFileType;
use crate::result::{SheetError, SheetResult};
use crate::xml::drawings::{Drawings};
use crate::xml::relationships::Relationships;
use crate::xml::worksheet::WorkSheet;
use crate::xml::workbook::Workbook;
use crate::xml::style::StyleSheet;

#[derive(Debug)]
pub struct Sheet {
    pub(crate) id: u32,
    pub(crate) name: String,
    workbook: Rc<RefCell<Workbook>>,
    worksheet: WorkSheet,
    worksheet_rel: Relationships,
    style_sheet: Rc<RefCell<StyleSheet>>,
    content_types: Rc<RefCell<xml::content_types::ContentTypes>>,
    medias: Rc<RefCell<xml::medias::Medias>>,
    drawings: HashMap<u32, Drawings>,
    drawings_rel: HashMap<u32, Relationships>,
}

impl Write for Sheet {}
impl Row for Sheet {}
impl Col for Sheet {}

impl Sheet {
    pub(crate) fn save_as<P: AsRef<Path>>(&mut self, file_path: P) -> SheetResult<()> {
        self.worksheet.save(&file_path, self.id);
        self.worksheet_rel.save(&file_path, XlsxFileType::WorksheetRels(self.id));
        self.drawings.iter_mut().for_each(|(id, d)| d.save(&file_path, *id));
        self.drawings_rel.iter_mut().for_each(|(id, d)| d.save(&file_path, XlsxFileType::DrawingRels(*id)));
        Ok(())
    }
}

impl Sheet {
    pub fn max_column(&self) -> u32 {
        let worksheet = &self.worksheet;
        let sheet_data = & worksheet.sheet_data;
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

    pub fn set_selection<L: Location>(&mut self, loc: L) {
        let worksheet = &mut self.worksheet;
        let sheet_views = &mut worksheet.sheet_views;
        sheet_views.set_selection(loc.to_ref());
    }

    pub fn freeze_panes<L: Location>(&mut self, loc: L) {
        let worksheet = &mut self.worksheet;
        let sheet_views = &mut worksheet.sheet_views;
        sheet_views.set_freeze_panes(loc);
    }

    pub fn set_default_row(&mut self, height: f64) {
        let worksheet = &mut self.worksheet;
        worksheet.set_default_row_height(height);
    }

    pub fn hide_unused_rows(&mut self, hide: bool) {
        let worksheet = &mut self.worksheet;
        worksheet.hide_unused_rows(hide);
    }

    pub fn outline_settings(& mut self, visible: bool, symbols_below: bool, symbols_right: bool, auto_style: bool) {
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
        let worksheet = &mut self.worksheet;
        worksheet.set_tab_color(tab_color);
    }

    pub fn set_background<P: AsRef<Path>>(&mut self, filename: &P) {
        let r_id = self.add_background(filename);
        let worksheet = &mut self.worksheet;
        worksheet.set_background(r_id);
    }

    pub fn insert_image<L: Location, P: AsRef<Path>>(&mut self, loc: L, filename: &P) {
        let (from_row, from_col) = loc.to_location();
        let (to_row, to_col) = (5 + from_row, 5 + from_col);
        let r_id = self.add_drawing_image((from_row, from_col, to_row, to_col), filename);
        let worksheet = &mut self.worksheet;
        worksheet.insert_image(loc, r_id);
    }

    pub fn id(&self) -> u32 {
        self.id
    }
}

impl Sheet {
    pub(crate) fn add_drawings(&mut self, tmp_path: &str) {
        let worksheet_rel = &mut self.worksheet_rel;
        let drawings_id = worksheet_rel.list_drawings();
        drawings_id.iter().for_each(|&id| {
            self.drawings.insert(id, Drawings::from_path(tmp_path, id).unwrap());
        });
    }

    pub(crate) fn from_xml<P: AsRef<Path>>(
        sheet_id: u32,
        name: &str,
        tmp_path: P,
        workbook: Rc<RefCell<Workbook>>,
        // worksheets_rel: Rc<RefCell<HashMap<u32, Relationships>>>,
        style_sheet: Rc<RefCell<StyleSheet>>,
        content_types: Rc<RefCell<xml::content_types::ContentTypes>>,
        medias: Rc<RefCell<xml::medias::Medias>>,
    ) -> Sheet {
        let worksheet = WorkSheet::from_path(&tmp_path, sheet_id).unwrap_or_default();
        let worksheet_rel = Relationships::from_path(&tmp_path, XlsxFileType::WorksheetRels(sheet_id)).unwrap_or_default();
        let mut drawings = HashMap::new();
        let mut drawings_rel = HashMap::new();
        worksheet_rel.list_drawings().iter().for_each(|&id| {
            drawings.insert(id, Drawings::from_path(&tmp_path, id).unwrap());
            drawings_rel.insert(id, Relationships::from_path(&tmp_path, XlsxFileType::DrawingRels(id)).unwrap());
        });
        Sheet {
            id: sheet_id,
            name: String::from(name),
            workbook,
            worksheet,
            worksheet_rel,
            style_sheet,
            content_types,
            medias,
            drawings,
            drawings_rel,
        }
    }
}