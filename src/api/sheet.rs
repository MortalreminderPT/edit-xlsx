use std::cell::RefCell;
use std::rc::{Rc, Weak};
use crate::api::format::Format;
use crate::result::SheetResult;
use crate::xml::manage::{Borrow, Create, XmlManager};
use crate::xml::sheet_data::cell_values::{CellDisplay, CellType};

#[derive(Debug)]
pub struct Sheet {
    pub(crate) id: u32,
    pub(crate) name: String,
    xml_manager: Rc<RefCell<XmlManager>>,
}

/// style
impl Sheet {
    fn add_format(&mut self, format: &Format) -> u32 {
        let mut binding = &mut self.xml_manager.borrow_mut();
        let style_sheet = binding.borrow_style_sheet_mut();
        style_sheet.add_format(format)
    }
}

impl Sheet {
    fn write_all<T: CellDisplay + CellType>(&mut self, row_id: u32, col_id: u32, text_id: T, style_id: Option<u32>) -> SheetResult<()> {
        let mut binding = &mut self.xml_manager.borrow_mut();
        let sheet_xml = binding.borrow_worksheet_mut(self.id);
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
        let binding = &mut self.xml_manager.borrow_mut();
        let sheet_xml = binding.borrow_worksheet_mut(self.id);
        let sheet_data = &mut sheet_xml.sheet_data;
        sheet_data.update_or_create_row(row, height, None)?;
        Ok(())
    }

    pub fn set_row_with_format(&mut self, row: u32, height: f64, format: &Format) {
        let style_id: u32 = self.add_format(format);
        let binding = &mut self.xml_manager.borrow_mut();
        let sheet_xml = binding.borrow_worksheet_mut(self.id);
        let sheet_data = &mut sheet_xml.sheet_data;
        sheet_data.update_or_create_row(row, height, Some(style_id));
    }

    pub fn set_row_pixels(&mut self, row: u32, height: f64) {
        self.set_row(row, 0.5 * height);
    }

    pub fn set_row_pixels_with_format(&mut self, row: u32, height: f64, format: &Format) {
        self.set_row_with_format(row, 0.5 * height, format)
    }

    pub fn set_column(&mut self, first_col: u32, last_col: u32, width: f64) {
        let binding = &mut self.xml_manager.borrow_mut();
        let sheet_xml = binding.borrow_worksheet_mut(self.id);
        sheet_xml.create_col(first_col, last_col, Some(width), None, 0);
    }

    pub fn set_column_with_format(&mut self, first_col: u32, last_col: u32, width: f64, format: &Format) {
        let style_id: u32 = self.add_format(format);
        let binding = &mut self.xml_manager.borrow_mut();
        let sheet_xml = binding.borrow_worksheet_mut(self.id);
        let col = sheet_xml.create_col(first_col, last_col, Some(width), Some(style_id), 0);
    }

    pub fn set_column_pixels(&mut self, first_col: u32, last_col: u32, width: f64) {
        self.set_column(first_col, last_col, 0.5 * width)
    }

    pub fn max_column(&mut self) -> u32 {
        let binding = &mut self.xml_manager.borrow_mut();
        let sheet_xml = binding.borrow_worksheet_mut(self.id);
        let sheet_data = &mut sheet_xml.sheet_data;
        sheet_data.max_col()
    }

    fn autofit(&mut self) {
        todo!();
        let binding = &mut self.xml_manager.borrow_mut();
        let sheet_xml = binding.borrow_worksheet_mut(self.id);
        sheet_xml.autofit_cols();
    }

    pub fn get_name(&self) -> &str {
        &self.name
    }

    pub fn activate(&self) {
        let binding = &mut self.xml_manager.borrow_mut();
        let mut book_views = &mut binding.borrow_workbook_mut().book_views;
        book_views.book_views[0].active_tab = Some(self.id - 1);
    }

    pub fn select(&self) {
        let binding = &mut self.xml_manager.borrow_mut();
        let mut sheet_views = &mut binding.borrow_worksheet_mut(self.id).sheet_views;
        sheet_views.sheet_view[0].tab_selected = Some(1);
    }

    pub fn hide(&self) {
        let binding = &mut self.xml_manager.borrow_mut();
        let sheet = &mut binding.borrow_workbook_mut()
            .sheets.sheets
            .iter_mut()
            .find(|s| { s.sheet_id == self.id })
            .unwrap();
        sheet.state = Some(String::from("hidden"));
    }

    pub fn set_first_sheet(&mut self) {
        self.change_id(1);
    }

    pub fn merge_range<T: CellDisplay + CellType>(&mut self, first_row: u32, first_col:u32, last_row:u32 , last_col:u32, data: T, format:&Format) -> SheetResult<()> {
        {
            let binding = &mut self.xml_manager.borrow_mut();
            let worksheet = &mut binding.borrow_worksheet_mut(self.id);
            worksheet.add_merge_cell(first_row, first_col, last_row, last_col);
        }
        self.write_with_format(first_row, first_col, data, format)?;
        Ok(())
    }

    fn change_id(&mut self, id: u32) {
        let loc = (self.id - 1) as usize;
        let binding = &mut self.xml_manager.borrow_mut();
        let sheets = &mut binding.borrow_workbook_mut()
            .sheets.sheets;
        sheets[loc].change_id(id);
        sheets[0].change_id(self.id);
        sheets.swap(0, loc);
    }

    pub fn id(&self) -> u32 {
        self.id
    }
}

impl Sheet {
    pub(crate) fn new(xml_manager: Rc<RefCell<XmlManager>>) -> Sheet {
        let (id, name) = xml_manager.borrow_mut().create_worksheet();
        Self::from_xml(id, &name, Rc::clone(&xml_manager))
    }

    pub(crate) fn from_xml(sheet_id: u32, name: &str, xml_manager: Rc<RefCell<XmlManager>>) -> Sheet {
        Sheet {
            id: sheet_id,
            name: String::from(name),
            xml_manager,
        }
    }
}