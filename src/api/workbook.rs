use std::fs;
use std::cell::RefCell;
use std::path::Path;
use std::rc::Rc;
use crate::api::sheet::Sheet;
use crate::utils::zip_util;
use crate::result::{SheetError, WorkbookError, WorkbookResult};
use crate::xml::manage::{Borrow, Create, XmlIo, XmlManager};

#[derive(Debug)]
pub struct Workbook {
    pub sheets: Vec<Sheet>,
    pub(crate) tmp_path: String,
    pub(crate) file_path: String,
    xml_manager: Rc<RefCell<XmlManager>>,
}

impl Workbook {
    pub fn new() -> Workbook {
        let mut wb = Self::from_path("resource/new.xlsx").unwrap();
        wb.file_path = String::from("./new.xlsx");
        wb
    }
    
    pub fn get_worksheet(&mut self, id: u32) -> WorkbookResult<&mut Sheet> {
        let sheet = self.sheets
            .iter_mut()
            .find(|sheet| sheet.id == id);
        match sheet {
            Some(sheet) => Ok(sheet),
            None => Err(WorkbookError::SheetError(SheetError::FileNotFound))
        }
    }

    pub fn add_worksheet(&mut self) -> WorkbookResult<&mut Sheet> {
        let (sheet_id, worksheet) = self.xml_manager.borrow_mut().create_worksheet();
        let sheet = Sheet::from_xml(sheet_id, Rc::clone(&self.xml_manager));
        self.sheets.push(sheet);
        self.get_worksheet(sheet_id)
    }
}

impl Workbook {
    pub fn from_path<P: AsRef<Path>>(file_path: P) -> WorkbookResult<Workbook> {
        let tmp_path = Workbook::extract_tmp_dir(&file_path)?;
        let xml_manager = XmlManager::from_path(&tmp_path)?;
        let xml_manager = Rc::new(RefCell::new(xml_manager));
        let sheets = xml_manager.borrow().borrow_workbook().sheets.sheets.iter().map(
            |sheet_xml| Sheet::from_xml(sheet_xml.sheet_id, Rc::clone(&xml_manager))
        ).collect();
        Ok(Workbook {
            xml_manager,
            sheets,
            tmp_path,
            file_path: file_path.as_ref().to_str().unwrap().to_string(),
        })
    }

    // fn create_tmp_dir<P: AsRef<Path>>(file_path: P) -> WorkbookResult<String> {
    //
    //     Ok(zip_util::extract_dir(file_path)?)
    // }

    fn extract_tmp_dir<P: AsRef<Path>>(file_path: P) -> WorkbookResult<String> {
        Ok(zip_util::extract_dir(file_path)?)
    }

    pub fn save_as<P: AsRef<Path>>(&mut self, file_path: P) -> WorkbookResult<()> {
        // save files
        self.xml_manager.borrow_mut().save(&self.tmp_path);
        // package files
        zip_util::zip_dir(&self.tmp_path, file_path)?;
        Ok(())
    }

    pub fn save(&mut self) -> WorkbookResult<()> {
        self.save_as(&self.file_path.clone())
    }
}

impl Drop for Workbook {
    fn drop(&mut self) {
        fs::remove_dir_all(&self.tmp_path);
    }
}