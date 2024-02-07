use std::cell::RefCell;
use std::rc::Rc;
use crate::api::format::Format;
use crate::result::SheetResult;
use crate::xml::manage::{Borrow, Create, EditRow, XmlManager};
use crate::xml::sheet_data::cell_values::{CellDisplay, CellType};

#[derive(Debug)]
pub struct Sheet {
    pub(crate) id: u32,
    xml_manager: Rc<RefCell<XmlManager>>,
}

impl Sheet {
    fn write_all<T: CellDisplay + CellType>(&mut self, row_id: u32, col_id: u32, text_id: T, style_id: Option<u32>) -> SheetResult<()> {
        let mut binding = &mut self.xml_manager.borrow_mut();
        let sheet_xml = binding.borrow_worksheet_mut(self.id);
        let sheet_data = &mut sheet_xml.sheet_data;
        let row = match sheet_data.get(row_id) {
            Ok(row) => row,
            Err(_) => sheet_data.create(row_id)?
        };
        row.update_or_create_cell(col_id, text_id, style_id);
        Ok(())
    }

    pub fn write<T: CellDisplay + CellType>(&mut self, row_id: u32, col_id: u32, text: T) -> SheetResult<()> {
        self.write_all(row_id, col_id, text, None)?;
        Ok(())
    }

    pub fn write_with_format<T: CellDisplay + CellType>(&mut self, row_id: u32, col_id: u32, text: T, format: &Format) -> SheetResult<()> {
        let mut style_id: Option<u32> = None;
        {
            let mut binding = &mut self.xml_manager.borrow_mut();
            let style_sheet = binding.borrow_style_sheet_mut();
            // todo get default format from cell
            style_id = Some(style_sheet.add_format(format));
        }
        self.write_all(row_id, col_id, text, style_id)?;
        Ok(())
    }

    pub fn id(&self) -> u32 {
        self.id
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