use std::cell::RefCell;
use std::collections::HashMap;
use std::path::Path;
use std::rc::Rc;
use crate::api::format::Format;
use crate::{FormatColor, WorkbookResult, xml};
use crate::result::{SheetError, SheetResult, WorkbookError};
use crate::xml::worksheet::WorkSheet;
use crate::xml::workbook::Workbook;
use crate::xml::sheet_data::cell_values::{CellDisplay, CellType};
use crate::xml::style::StyleSheet;
use crate::xml::worksheet_rel::Relationships;

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

/// style
impl Sheet {
    fn add_format(&mut self, format: &Format) -> u32 {
        let mut style_sheet = self.style_sheet.borrow_mut();
        style_sheet.add_format(format)
    }
}

impl Sheet {
    
    fn write_all<T: CellDisplay + CellType>(&mut self, row_id: u32, col_id: u32, text_id: T, style_id: Option<u32>) -> SheetResult<()> {
        let worksheets = &mut self.worksheets.borrow_mut();
        let sheet_xml = worksheets.get_mut(&self.id).unwrap();
        let sheet_data = &mut sheet_xml.sheet_data;
        let row = match sheet_data.get_row(row_id) {
            Ok(row) => row,
            Err(_) => sheet_data.create_row(row_id)
        };
        row.update_or_create_cell(col_id, text_id, style_id);
        Ok(())
    }

    pub fn write<T: CellDisplay + CellType>(&mut self, row_id: u32, col_id: u32, text: T) -> SheetResult<()> {
        self.write_all(row_id, col_id, text, None)?;
        Ok(())
    }

    pub fn write_with_format<T: CellDisplay + CellType>(&mut self, row_id: u32, col_id: u32, text: T, format: &Format) -> SheetResult<()> {
        let style_id: u32 = self.add_format(format);
        self.write_all(row_id, col_id, text, Some(style_id))?;
        Ok(())
    }

    pub fn set_row(&mut self, row: u32, height: f64) -> SheetResult<()> {
        let worksheets = &mut self.worksheets.borrow_mut();
        let sheet_xml = worksheets.get_mut(&self.id).unwrap();
        let sheet_data = &mut sheet_xml.sheet_data;
        sheet_data.update_or_create_row(row, height, None)?;
        Ok(())
    }

    pub fn set_row_with_format(&mut self, row: u32, height: f64, format: &Format) -> SheetResult<()> {
        let style_id: u32 = self.add_format(format);
        let worksheets = &mut self.worksheets.borrow_mut();
        let sheet_xml = worksheets.get_mut(&self.id).unwrap();
        let sheet_data = &mut sheet_xml.sheet_data;
        sheet_data.update_or_create_row(row, height, Some(style_id))?;
        Ok(())
    }

    pub fn set_row_pixels(&mut self, row: u32, height: f64) -> SheetResult<()> {
        self.set_row(row, 0.5 * height)
    }

    pub fn set_row_pixels_with_format(&mut self, row: u32, height: f64, format: &Format) -> SheetResult<()> {
        self.set_row_with_format(row, 0.5 * height, format)
    }

    pub fn set_column(&mut self, first_col: u32, last_col: u32, width: f64) -> SheetResult<()> {
        let worksheets = &mut self.worksheets.borrow_mut();
        let sheet_xml = worksheets.get_mut(&self.id).unwrap();
        sheet_xml.create_col(first_col, last_col, Some(width), None, 0)?;
        Ok(())
    }

    pub fn set_column_with_format(&mut self, first_col: u32, last_col: u32, width: f64, format: &Format) -> SheetResult<()> {
        let style_id = self.add_format(format);
        let worksheets = &mut self.worksheets.borrow_mut();
        let sheet_xml = worksheets.get_mut(&self.id).unwrap();
        sheet_xml.create_col(first_col, last_col, Some(width), Some(style_id), 0)?;
        Ok(())
    }

    pub fn set_column_pixels(&mut self, first_col: u32, last_col: u32, width: f64) -> SheetResult<()> {
        self.set_column(first_col, last_col, 0.5 * width)
    }

    pub fn max_column(&mut self) -> u32 {
        let worksheets = &mut self.worksheets.borrow_mut();
        let sheet_xml = worksheets.get_mut(&self.id).unwrap();
        let sheet_data = &mut sheet_xml.sheet_data;
        sheet_data.max_col()
    }

    fn autofit(&mut self) {
        todo!();
        let worksheets = &mut self.worksheets.borrow_mut();
        let sheet_xml = worksheets.get_mut(&self.id).unwrap();
        sheet_xml.autofit_cols();
    }

    pub fn get_name(&self) -> &str {
        &self.name
    }

    pub fn activate(&mut self) {
        let workbook = &mut self.workbook.borrow_mut();
        let book_views = &mut workbook.book_views;
        book_views.book_views[0].active_tab = Some(self.id - 1);
    }

    pub fn select(&mut self) {
        let worksheets = &mut self.worksheets.borrow_mut();
        let sheet_xml = worksheets.get_mut(&self.id).unwrap();
        let sheet_views = &mut sheet_xml.sheet_views;
        sheet_views.sheet_view[0].tab_selected = Some(1);
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

    pub fn set_first_sheet(&mut self) {
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
        Ok(())
    }

    pub fn merge_range<T: CellDisplay + CellType>(&mut self, first_row: u32, first_col:u32, last_row:u32 , last_col:u32, data: T, format:&Format) -> SheetResult<()> {
        {
            let worksheets = &mut self.worksheets.borrow_mut();
            let worksheet = worksheets.get_mut(&self.id).unwrap();
            worksheet.add_merge_cell(first_row, first_col, last_row, last_col);
        }
        self.write_with_format(first_row, first_col, data, format)?;
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
        let mut worksheets_rel = self.worksheets_rel.borrow_mut();
        if let None = worksheets_rel.get_mut(&self.id) {
            worksheets_rel.insert(self.id, Relationships::default());
        }
        let worksheet_rel = worksheets_rel.get_mut(&self.id).unwrap();
        self.content_types.borrow_mut().add_png();
        let image_id = self.medias.borrow_mut().add_media(filename);
        worksheet.set_background(image_id);
        worksheet_rel.add_image(image_id);
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