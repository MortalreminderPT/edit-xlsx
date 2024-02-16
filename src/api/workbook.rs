use std::{fs, slice};
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
        let (sheet_id, name) = self.xml_manager.borrow_mut().create_worksheet();
        let sheet = Sheet::from_xml(sheet_id, &name, Rc::clone(&self.xml_manager));
        self.sheets.push(sheet);
        self.get_worksheet(sheet_id)
    }

    pub fn add_worksheet_by_name(&mut self, name: &str) -> WorkbookResult<&mut Sheet> {
        let sheet_id = self.xml_manager.borrow_mut().create_worksheet_by_name(name);
        let sheet = Sheet::from_xml(sheet_id, name, Rc::clone(&self.xml_manager));
        self.sheets.push(sheet);
        self.get_worksheet(sheet_id)
    }

    pub fn set_size(&mut self, width: u32, height: u32) -> WorkbookResult<()> {
        let mut binding = self.xml_manager.borrow_mut();
        let book_view = binding.borrow_workbook_mut().book_views.book_views.get_mut(0).unwrap();
        book_view.window_width = width;
        book_view.window_height = height;
        Ok(())
    }

    pub fn set_tab_ratio(&mut self, tab_ratio: f64) -> WorkbookResult<()> {
        let tab_ratio = (tab_ratio * 10.0).round() as u32;
        let mut binding = self.xml_manager.borrow_mut();
        let book_view = binding.borrow_workbook_mut().book_views.book_views.get_mut(0).unwrap();
        book_view.tab_ratio = Some(tab_ratio);
        Ok(())
    }

    fn define_name(&mut self, name: &str, formula: &str) -> WorkbookResult<()> {
        let mut binding = self.xml_manager.borrow_mut();
        let book_view = binding.borrow_workbook_mut().book_views.book_views.get_mut(0).unwrap();
        Ok(())
    }

    pub fn worksheets(&mut self) -> slice::IterMut<Sheet> {
        self.sheets.iter_mut()
    }

    pub fn get_worksheet_by_name(&mut self, name: &str) -> WorkbookResult<&mut Sheet> {
        let sheet = self.sheets
            .iter_mut()
            .find(|sheet| sheet.name == name);
        match sheet {
            Some(sheet) => Ok(sheet),
            None => Err(WorkbookError::SheetError(SheetError::FileNotFound))
        }
    }
}

impl Workbook {
    pub fn from_path<P: AsRef<Path>>(file_path: P) -> WorkbookResult<Workbook> {
        let tmp_path = Workbook::extract_tmp_dir(&file_path)?;
        let xml_manager = XmlManager::from_path(&tmp_path)?;
        let xml_manager = Rc::new(RefCell::new(xml_manager));
        let sheets = xml_manager.borrow().borrow_workbook().sheets.sheets.iter().map(
            |sheet_xml| Sheet::from_xml(sheet_xml.sheet_id, &sheet_xml.name, Rc::clone(&xml_manager))
        ).collect();
        Ok(Workbook {
            xml_manager,
            sheets,
            tmp_path,
            file_path: file_path.as_ref().to_str().unwrap().to_string(),
        })
    }
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
        fs::remove_dir_all(&self.tmp_path).unwrap();
    }
}