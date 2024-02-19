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
use crate::api::location::Location;
use crate::api::worksheet::col::Col;
use crate::api::worksheet::image::_Image;
use crate::api::worksheet::row::Row;
use crate::api::worksheet::write::Write;
use crate::result::SheetError;
use crate::xml::worksheet::WorkSheet;
use crate::xml::workbook::Workbook;
use crate::xml::style::StyleSheet;

#[derive(Debug)]
pub struct Sheet {
    pub(crate) id: u32,
    pub(crate) name: String,
    workbook: Rc<RefCell<Workbook>>,
    worksheets: Rc<RefCell<HashMap<u32, WorkSheet>>>,
    worksheets_rel: Rc<RefCell<HashMap<u32, xml::worksheet_rel::Relationships>>>,
    style_sheet: Rc<RefCell<StyleSheet>>,
    content_types: Rc<RefCell<xml::content_types::ContentTypes>>,
    medias: Rc<RefCell<xml::medias::Medias>>
}

impl Write for Sheet {}
impl Row for Sheet {}
impl Col for Sheet {}

impl Sheet {
    pub fn max_column(&mut self) -> u32 {
        let worksheets = &mut self.worksheets.borrow_mut();
        let worksheet = worksheets.get_mut(&self.id).unwrap();
        let sheet_data = &mut worksheet.sheet_data;
        sheet_data.max_col()
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
        let worksheets = &mut self.worksheets.borrow_mut();
        let worksheet = worksheets.get_mut(&self.id).unwrap();
        let sheet_views = &mut worksheet.sheet_views;
        sheet_views.set_tab_selected(1);
    }

    pub fn set_top_left_cell<L: Location>(&mut self, loc: L) {
        let worksheets = &mut self.worksheets.borrow_mut();
        let worksheet = worksheets.get_mut(&self.id).unwrap();
        let sheet_views = &mut worksheet.sheet_views;
        sheet_views.set_top_left_cell(&loc.to_ref());
    }

    pub fn set_zoom(&mut self, zoom_scale: u16) {
        let worksheets = &mut self.worksheets.borrow_mut();
        let worksheet = worksheets.get_mut(&self.id).unwrap();
        let sheet_views = &mut worksheet.sheet_views;
        sheet_views.set_zoom_scale(zoom_scale);
    }

    pub fn set_selection<L: Location>(&mut self, loc: L) {
        let worksheets = &mut self.worksheets.borrow_mut();
        let worksheet = worksheets.get_mut(&self.id).unwrap();
        let sheet_views = &mut worksheet.sheet_views;
        sheet_views.set_selection(loc.to_ref());
    }

    pub fn freeze_panes<L: Location>(&mut self, loc: L) {
        let worksheets = &mut self.worksheets.borrow_mut();
        let worksheet = worksheets.get_mut(&self.id).unwrap();
        let sheet_views = &mut worksheet.sheet_views;
        let from = loc.to_location();
        let to = (1 + from.0, 1 + from.1);
        sheet_views.set_freeze_panes(&from.to_ref(), &to.to_ref());
    }

    pub fn set_default_row(&mut self, height: f64) {
        let worksheets = &mut self.worksheets.borrow_mut();
        let worksheet = worksheets.get_mut(&self.id).unwrap();
        worksheet.sheet_format_pr.default_row_height = height;
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
        let sheet_target = self.worksheets.borrow_mut().remove(&id).ok_or(SheetError::FileNotFound)?;
        let sheet = self.worksheets.borrow_mut().remove(&self.id).ok_or(SheetError::FileNotFound)?;
        self.worksheets.borrow_mut().insert(id, sheet);
        self.worksheets.borrow_mut().insert(self.id, sheet_target);
        // swap sheet rel
        let sheet_rel_target = self.worksheets_rel.borrow_mut().remove(&id);
        let sheet_rel = self.worksheets_rel.borrow_mut().remove(&self.id);
        if let Some(sheet_rel_target) = sheet_rel_target {
            self.worksheets_rel.borrow_mut().insert(self.id, sheet_rel_target);
        }
        if let Some(sheet_rel) = sheet_rel {
            self.worksheets_rel.borrow_mut().insert(id, sheet_rel);
        }
        // swap workbook sheet
        let loc = (self.id - 1) as usize;
        let workbook = &mut self.workbook.borrow_mut();
        let sheets = &mut workbook.sheets.sheets;
        sheets[loc].change_id(id);
        sheets[0].change_id(self.id);
        sheets.swap(0, loc);
        // change self id

        Ok(())
    }

    pub fn set_tab_color(&mut self, tab_color: &FormatColor) {
        let worksheets = &mut self.worksheets.borrow_mut();
        let worksheet = worksheets.get_mut(&self.id).unwrap();
        worksheet.set_tab_color(tab_color);
    }

    pub fn set_background<P: AsRef<Path>>(&mut self, filename: &P) {
        let worksheets = &mut self.worksheets.borrow_mut();
        let worksheet = worksheets.get_mut(&self.id).unwrap();
        let r_id = self.add_image(filename);
        worksheet.set_background(r_id);
    }

    pub fn id(&self) -> u32 {
        self.id
    }
}

impl Sheet {
    pub(crate) fn from_xml(
        sheet_id: u32,
        name: &str,
        workbook: Rc<RefCell<Workbook>>,
        worksheets: Rc<RefCell<HashMap<u32, WorkSheet>>>,
        worksheets_rel: Rc<RefCell<HashMap<u32, xml::worksheet_rel::Relationships>>>,
        style_sheet: Rc<RefCell<StyleSheet>>,
        content_types: Rc<RefCell<xml::content_types::ContentTypes>>,
        medias: Rc<RefCell<xml::medias::Medias>>,
    ) -> Sheet {
        Sheet {
            id: sheet_id,
            name: String::from(name),
            workbook,
            worksheets,
            worksheets_rel,
            style_sheet,
            content_types,
            medias,
        }
    }
}