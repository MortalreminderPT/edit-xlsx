use std::cell::RefCell;
use std::rc::Rc;
use crate::api::format::Format;
use crate::result::SheetResult;
use crate::xml::manage::{Borrow, Create, EditRow, XmlManager};

#[derive(Debug)]
pub struct Sheet {
    pub(crate) id: u32,
    xml_manager: Rc<RefCell<XmlManager>>,
}

impl Sheet {
    fn write_with_id(&mut self, row_id: u32, col_id: u32, text_id: Option<u32>, style_id: Option<u32>) -> SheetResult<()> {
        let mut binding = &mut self.xml_manager.borrow_mut();
        let sheet_xml = binding.borrow_worksheet_mut(self.id);
        let sheet_data = &mut sheet_xml.sheet_data;
        let row = match sheet_data.get(row_id) {
            Ok(row) => row,
            Err(_) => sheet_data.create(row_id)?
        };
        match row.get_cell(col_id) {
            Some(cell) => {
                cell.text = text_id;
                cell.style = style_id;
                cell.text_type = Some("s".to_string());
            },
            None => { row.create_cell(row_id, col_id, text_id, style_id, "s"); },
        };
        Ok(())
    }

    pub fn write(&mut self, row_id: u32, col_id: u32, text: &str) -> SheetResult<()> {
        let mut text_id: Option<u32> = None;
        {
            let mut binding = &mut self.xml_manager.borrow_mut();
            let shared_string = binding.borrow_shared_string_mut();
            text_id = Some(shared_string.add_text(text));
        }
        self.write_with_id(row_id, col_id, text_id, None)?;
        Ok(())
    }

    pub fn write_with_format(&mut self, row_id: u32, col_id: u32, text: &str, format: Format) -> SheetResult<()> {
        let mut text_id: Option<u32> = None;
        let mut style_id: Option<u32> = None;
        {
            let mut binding = &mut self.xml_manager.borrow_mut();
            let shared_string = binding.borrow_shared_string_mut();
            text_id = Some(shared_string.add_text(text));
            let style_sheet = binding.borrow_style_sheet_mut();
            style_id = Some(style_sheet.add_format(format));
        }
        self.write_with_id(row_id, col_id, text_id, style_id)?;
        Ok(())
    }
}

impl Sheet {
    pub(crate) fn new(xml_manager: Rc<RefCell<XmlManager>>) -> Sheet {
        let (id, worksheet) = xml_manager.borrow_mut().create_worksheet();
        Self::from_xml(id, Rc::clone(&xml_manager))
    }

    pub(crate) fn from_xml(sheet_id: u32, xml_manager: Rc<RefCell<XmlManager>>) -> Sheet {
        Sheet {
            id: sheet_id,
            xml_manager,
        }
    }
}