mod tests;
mod rel;

use std::fs;
use std::cell::RefCell;
use std::io::{Read, Seek, Write};
use std::path::Path;
use std::rc::Rc;
use crate::sheet::Sheet;
use crate::utils::zip_util;
use crate::result::WorkbookResult;
use crate::xml::facade::{Borrow, XmlIo, XmlManager};

#[derive(Debug)]
pub struct Workbook {
    pub sheets: Vec<Sheet>,
    pub(crate) tmp_path: String,
    pub(crate) file_path: String,
    xml_manager: Rc<RefCell<XmlManager>>,
}

impl Workbook {
    pub fn get_sheet_mut(&mut self, id: u32) -> Option<&mut Sheet> {
        self.sheets.iter_mut().find(|sheet| sheet.id == id)
    }

    pub fn get_sheet(&self, id: u32) -> Option<&Sheet> {
        self.sheets.iter().find(|sheet| sheet.id == id)
    }

    pub fn add_worksheet(&mut self) -> Option<&mut Sheet> {
        let max_id = self.sheets.iter().map(|sheet| sheet.id).max().unwrap();
        let sheet = Sheet::new(max_id + 1, Rc::clone(&self.xml_manager));
        self.sheets.push(sheet);
        self.get_sheet_mut(max_id + 1)
    }
}

impl Workbook {
    pub fn from_path<P: AsRef<Path>>(file_path: P) -> Workbook {
        let tmp_path = Workbook::create_tmp_dir(&file_path).unwrap();
        let xml_manager = XmlManager::from_path(&tmp_path);
        let xml_manager = Rc::new(RefCell::new(xml_manager));
        let sheets = xml_manager.borrow().borrow_workbook().sheets.sheets.iter().map(
            |sheet_xml| Sheet::from_xml(sheet_xml.sheet_id, Rc::clone(&xml_manager))
        ).collect();
        Workbook {
            xml_manager,
            sheets,
            tmp_path,
            file_path: file_path.as_ref().to_str().unwrap().to_string(),
        }
    }

    fn create_tmp_dir<P: AsRef<Path>>(file_path: P) -> WorkbookResult<String> {
        Ok(zip_util::extract_dir(file_path)?)
    }

    pub fn save_as<P: AsRef<Path>>(&self, file_path: P) -> WorkbookResult<()> {
        // save files
        self.xml_manager.borrow_mut().save(&self.tmp_path);
        // package files
        zip_util::zip_dir(&self.tmp_path, file_path)?;
        Ok(())
    }

    pub fn save(&self) -> WorkbookResult<()> {
        self.save_as(&self.file_path)
    }
}

impl Drop for Workbook {
    fn drop(&mut self) {
        fs::remove_dir_all(&self.tmp_path).unwrap();
    }
}