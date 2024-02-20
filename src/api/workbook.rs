use std::{fs, slice};
use std::cell::RefCell;
use std::collections::HashMap;
use std::path::Path;
use std::rc::Rc;
use crate::api::worksheet::Sheet;
use crate::file::XlsxFileType;
use crate::utils::zip_util;
use crate::result::{SheetError, WorkbookError, WorkbookResult};
use crate::xml;
use crate::xml::content_types::ContentTypes;
use crate::xml::io::Io;
use crate::xml::medias::Medias;
use crate::xml::style::StyleSheet;
use crate::xml::relationships::Relationships;
use crate::xml::worksheet::WorkSheet;

#[derive(Debug)]
pub struct Workbook {
    pub sheets: Vec<Sheet>,
    pub(crate) tmp_path: String,
    pub(crate) file_path: String,
    closed: bool,
    workbook: Rc<RefCell<xml::workbook::Workbook>>,
    style_sheet: Rc<RefCell<StyleSheet>>,
    workbook_rel: Rc<RefCell<Relationships>>,
    content_types: Rc<RefCell<ContentTypes>>,
    medias: Rc<RefCell<Medias>>
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
            .find(|sheet| sheet.id == id).ok_or(SheetError::FileNotFound)?;
        Ok(sheet)
    }

    pub fn add_worksheet(&mut self) -> WorkbookResult<&mut Sheet> {
        let r_id = self.workbook_rel.borrow().next_id();
        let (id, name) = self.workbook.borrow_mut().add_worksheet(r_id)?;
        self.workbook_rel.borrow_mut().add_worksheet(r_id, id);
        // self.worksheets.borrow_mut().insert(id, WorkSheet::new());
        let sheet = Sheet::from_xml(
            id,
            &name,
            &self.tmp_path,
            Rc::clone(&self.workbook),
            // Rc::clone(&self.worksheets_rel),
            Rc::clone(&self.style_sheet),
            Rc::clone(&self.content_types),
            Rc::clone(&self.medias),
        );
        self.sheets.push(sheet);
        self.get_worksheet(id)
    }

    pub fn add_worksheet_by_name(&mut self, name: &str) -> WorkbookResult<&mut Sheet> {
        let r_id = self.workbook_rel.borrow().next_id();
        let id = self.workbook.borrow_mut().add_worksheet_by_name(r_id, name)?;
        self.workbook_rel.borrow_mut().add_worksheet(r_id, id);
        // self.worksheets.borrow_mut().insert(id, WorkSheet::new());
        let sheet = Sheet::from_xml(
            id,
            name,
            &self.tmp_path,
            Rc::clone(&self.workbook),
            // Rc::clone(&self.worksheets_rel),
            Rc::clone(&self.style_sheet),
            Rc::clone(&self.content_types),
            Rc::clone(&self.medias),
        );
        self.sheets.push(sheet);
        self.get_worksheet(id)
    }

    pub fn set_size(&mut self, width: u32, height: u32) -> WorkbookResult<()> {
        let workbook = &mut self.workbook.borrow_mut();
        let book_view = workbook.book_views.book_views.get_mut(0).unwrap();
        book_view.window_width = width;
        book_view.window_height = height;
        Ok(())
    }

    pub fn set_tab_ratio(&mut self, tab_ratio: f64) -> WorkbookResult<()> {
        let tab_ratio = (tab_ratio * 10.0).round() as u32;
        let workbook = &mut self.workbook.borrow_mut();
        let book_view = workbook.book_views.book_views.get_mut(0).unwrap();
        book_view.tab_ratio = Some(tab_ratio);
        Ok(())
    }

    // fn define_name(&mut self, name: &str, formula: &str) -> WorkbookResult<()> {
    //     let book_view = self.workbook.borrow_mut().book_views.book_views.get_mut(0).unwrap();
    //     Ok(())
    // }

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

    pub fn read_only_recommended(&mut self) -> WorkbookResult<()> {
        let workbook = &mut self.workbook.borrow_mut();
        let mut file_sharing = workbook.file_sharing.take().unwrap_or_default();
        file_sharing.read_only_recommended = 1;
        workbook.file_sharing = Some(file_sharing);
        Ok(())
    }
}

impl Workbook {
    pub fn from_path<P: AsRef<Path>>(file_path: P) -> WorkbookResult<Workbook> {
        let tmp_path = Workbook::extract_tmp_dir(&file_path)?;
        let workbook = xml::workbook::Workbook::from_path(&tmp_path)?;
        let workbook_rel = Relationships::from_path(&tmp_path, XlsxFileType::WorkbookRels)?;
        let style_sheet = StyleSheet::from_path(&tmp_path)?;
        let content_types = ContentTypes::from_path(&tmp_path)?;
        let medias = Medias::from_path(&tmp_path)?;
        let workbook = Rc::new(RefCell::new(workbook));
        let workbook_rel = Rc::new(RefCell::new(workbook_rel));
        let style_sheet = Rc::new(RefCell::new(style_sheet));
        let content_types = Rc::new(RefCell::new(content_types));
        let medias = Rc::new(RefCell::new(medias));

        let sheets = workbook.borrow().sheets.sheets.iter().map(
            |sheet_xml| {
                Sheet::from_xml(
                    sheet_xml.sheet_id,
                    &sheet_xml.name,
                    &tmp_path,
                    Rc::clone(&workbook),
                    Rc::clone(&style_sheet),
                    Rc::clone(&content_types),
                    Rc::clone(&medias),
                )
            }).collect::<Vec<Sheet>>();
        Ok(Workbook {
            sheets,
            tmp_path,
            file_path: file_path.as_ref().to_str().unwrap().to_string(),
            closed: false,
            workbook: Rc::clone(&workbook),
            workbook_rel: Rc::clone(&workbook_rel),
            style_sheet: Rc::clone(&style_sheet),
            content_types: Rc::clone(&content_types),
            medias: Rc::clone(&medias),
        })
    }

    fn extract_tmp_dir<P: AsRef<Path>>(file_path: P) -> WorkbookResult<String> {
        Ok(zip_util::extract_dir(file_path)?)
    }

    pub fn save_as<P: AsRef<Path>>(&mut self, file_path: P) -> WorkbookResult<()> {
        // save sheets
        self.sheets.iter_mut().for_each(|s|s.save_as(&self.tmp_path).unwrap());
        // save files
        self.workbook.borrow_mut().save(&self.tmp_path);
        self.style_sheet.borrow_mut().save(&self.tmp_path);
        self.workbook_rel.borrow_mut().save(&self.tmp_path, XlsxFileType::WorkbookRels);
        self.content_types.borrow_mut().save(&self.tmp_path);
        self.medias.borrow_mut().save(&self.tmp_path);
        // package files
        zip_util::zip_dir(&self.tmp_path, file_path)?;
        Ok(())
    }

    pub fn save(&mut self) -> WorkbookResult<()> {
        self.save_as(&self.file_path.clone())
    }

    pub fn finish(&mut self) {
        if !self.closed {
            fs::remove_dir_all(&self.tmp_path).unwrap();
            self.closed = true;
        }
    }
}

impl Drop for Workbook {
    fn drop(&mut self) {
        self.finish();
    }
}