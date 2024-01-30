use std::cell::RefCell;
use std::path::Path;
use std::rc::Rc;
use crate::sheet::xml::WorkSheet;
use crate::shared_string::SharedString;
use crate::sheet::result::SheetResult;

mod xml;
mod result;

#[derive(Debug)]
pub struct Sheet {
    id: u32,
    worksheet: WorkSheet,
    pub(crate) shared_string: Rc<RefCell<SharedString>>,
}

impl Sheet {
    pub fn write(&mut self, row_id: u32, col_id: u32, text: &str) -> SheetResult<()> {
        let mut row = match self.worksheet.sheet_data.get_row(row_id) {
            Some(row) => row,
            None => self.worksheet.sheet_data.create_row(row_id)
        };
        let mut cell = match row.get_cell(col_id) {
            Some(cell) => cell,
            None => row.create_cell(row_id, col_id),
        };
        let id = &self.shared_string.borrow_mut().add_text(text);
        cell.text = Some(id.to_string());
        Ok(())
    }
}

impl Sheet {
    pub(crate) fn from_path<P: AsRef<Path>>(file_path: P, sheet_id: u32, shared_string: Rc<RefCell<SharedString>>) -> Sheet {
        Sheet {
            id: sheet_id,
            worksheet: WorkSheet::from_path(file_path, sheet_id),
            shared_string,
        }
    }

    pub fn save<P: AsRef<Path>>(&self, file_path: P) {
        self.worksheet.save(file_path, self.id);
    }
}