use std::cell::RefCell;
use std::rc::Rc;
use crate::result::SheetResult;
use crate::xml::facade::{Borrow, Create, EditRow, XmlManager};

#[derive(Debug)]
pub struct Sheet {
    pub(crate) id: u32,
    xml_manager: Rc<RefCell<XmlManager>>,
}

impl Sheet {
    pub fn write(&mut self, row_id: u32, col_id: u32, text: &str) -> SheetResult<()> {
        let mut binding = &mut self.xml_manager.borrow_mut();
        let shared_string = binding.borrow_shared_string_mut();
        let text_id = shared_string.add_text(text);
        let sheet_xml = binding.borrow_worksheet_mut(self.id);
        let sheet_data = &mut sheet_xml.sheet_data;
        let row = match sheet_data.get(row_id) {
            Ok(row) => row,
            Err(_) => sheet_data.create(row_id)?
        };
        let cell = match row.get_cell(col_id) {
            Some(cell) => cell,
            None => row.create_cell(row_id, col_id)
        };
        cell.text = Some(text_id.to_string());
        Ok(())
    }
}

impl Sheet {
    pub(crate) fn new(sheet_id: u32, xml_manager: Rc<RefCell<XmlManager>>) -> Sheet {
        xml_manager.borrow_mut().create_worksheet(sheet_id);
        Self::from_xml(sheet_id, Rc::clone(&xml_manager))
    }

    pub(crate) fn from_xml(sheet_id: u32, xml_manager: Rc<RefCell<XmlManager>>) -> Sheet {
        Sheet {
            id: sheet_id,
            xml_manager,
        }
    }
}