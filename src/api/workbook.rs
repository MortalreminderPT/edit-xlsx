use std::{fs, slice};
use std::cell::RefCell;
use std::fs::File;
use std::hash::{Hash, Hasher};
use std::path::Path;
use std::rc::Rc;
use futures::executor::block_on;
use futures::join;
use zip::result::ZipError;
use crate::api::worksheet::WorkSheet;
use crate::file::XlsxFileType;
use crate::utils::{id_util, zip_util};
use crate::result::{WorkSheetError, WorkbookError, WorkbookResult};
use crate::{Properties, xml};
use crate::xml::content_types::ContentTypes;
use crate::xml::core_properties::CoreProperties;
use crate::xml::app_properties::AppProperties;
use crate::xml::io::{Io, IoV2};
use crate::xml::medias::Medias;
use crate::xml::metadata::Metadata;
use crate::xml::style::StyleSheet;
use crate::xml::relationships::Relationships;
use crate::xml::shared_string::SharedString;

#[derive(Debug)]
pub struct Workbook {
    pub sheets: Vec<WorkSheet>,
    pub(crate) tmp_path: String,
    pub(crate) file_path: String,
    closed: bool,
    pub(crate) workbook: Rc<RefCell<xml::workbook::Workbook>>,
    pub(crate) style_sheet: Rc<RefCell<StyleSheet>>,
    pub(crate) workbook_rel: Rc<RefCell<Relationships>>,
    pub(crate) content_types: Rc<RefCell<ContentTypes>>,
    pub(crate) medias: Rc<RefCell<Medias>>,
    pub(crate) metadata: Rc<RefCell<Metadata>>,
    pub(crate) core_properties: Option<CoreProperties>,
    pub(crate) app_properties: Option<AppProperties>,
    pub(crate) shared_string: Rc<SharedString>,
}

///
/// Private methods
///
impl Workbook {
    fn get_core_properties(&mut self) -> &mut CoreProperties {
        self.core_properties.get_or_insert(CoreProperties::from_path(&self.file_path).unwrap())
    }

    fn get_app_properties(&mut self) -> &mut AppProperties {
        self.app_properties.get_or_insert(AppProperties::from_path(&self.file_path).unwrap())
    }
}

impl Workbook {
    pub fn new() -> Workbook {
        let file_path = concat!(env!("CARGO_MANIFEST_DIR"), "/resources/new.xlsx");
        let mut wb = Self::from_path(file_path).unwrap();
        wb.file_path = String::from(file_path);
        wb
    }

    pub fn get_worksheet_mut(&mut self, id: u32) -> WorkbookResult<&mut WorkSheet> {
        let sheet = self.sheets
            .iter_mut()
            .find(|sheet| sheet.id == id).ok_or(WorkSheetError::FileNotFound)?;
        Ok(sheet)
    }

    pub fn get_worksheet(&self, id: u32) -> WorkbookResult<&WorkSheet> {
        let sheet = self.sheets
            .iter()
            .find(|sheet| sheet.id == id).ok_or(WorkSheetError::FileNotFound)?;
        Ok(sheet)
    }
    
    pub fn get_worksheet_by_name(&self, name: &str) -> WorkbookResult<&WorkSheet> {
        let sheet = self.sheets
            .iter()
            .find(|sheet| sheet.name == name);
        match sheet {
            Some(sheet) => Ok(sheet),
            None => Err(WorkbookError::SheetError(WorkSheetError::FileNotFound))
        }
    }
    
    pub fn get_worksheet_mut_by_name(&mut self, name: &str) -> WorkbookResult<& mut WorkSheet> {
        let sheet = self.sheets
            .iter_mut()
            .find(|sheet| sheet.name == name);
        match sheet {
            Some(sheet) => Ok(sheet),
            None => Err(WorkbookError::SheetError(WorkSheetError::FileNotFound))
        }
    }

    pub fn add_worksheet(&mut self) -> WorkbookResult<&mut WorkSheet> {
        let (r_id, target_id) = self.workbook_rel.borrow_mut().add_worksheet_v2();
        let (sheet_id, name) = self.workbook.borrow_mut().add_worksheet_v2(r_id, None)?;
        let worksheet = WorkSheet::add_worksheet(sheet_id, &name, target_id, self);
        self.sheets.push(worksheet);
        self.get_worksheet_mut(sheet_id)
    }

    pub fn add_worksheet_by_name(&mut self, name: &str) -> WorkbookResult<&mut WorkSheet> {
        let (r_id, target_id) = self.workbook_rel.borrow_mut().add_worksheet_v2();
        let (sheet_id, name) = self.workbook.borrow_mut().add_worksheet_v2(r_id, Some(name))?;
        let worksheet = WorkSheet::add_worksheet(sheet_id, &name, target_id, self);
        self.sheets.push(worksheet);
        self.get_worksheet_mut(sheet_id)
    }

    pub fn duplicate_worksheet(&mut self, id: u32) -> WorkbookResult<&mut WorkSheet> {
        let copy_worksheet = self.sheets
            .iter()
            .find(|sheet| sheet.id == id).ok_or(WorkSheetError::FileNotFound)?;
        let (r_id, target_id) = self.workbook_rel.borrow_mut().add_worksheet_v2();
        let (sheet_id, new_name) = self.workbook.borrow_mut().add_worksheet_v2(r_id, None)?;
        let worksheet = WorkSheet::from_worksheet(sheet_id, &new_name, target_id, copy_worksheet);
        self.sheets.push(worksheet);
        self.get_worksheet_mut(sheet_id)
    }

    pub fn duplicate_worksheet_by_name(&mut self, name: &str) -> WorkbookResult<&mut WorkSheet> {
        let copy_worksheet = self.sheets
            .iter()
            .find(|sheet| sheet.name == name).ok_or(WorkSheetError::FileNotFound)?;
        let new_name = format!("{} Duplicated", name);
        let (r_id, target_id) = self.workbook_rel.borrow_mut().add_worksheet_v2();
        let (sheet_id, _) = self.workbook.borrow_mut().add_worksheet_v2(r_id, Some(&new_name))?;
        let worksheet = WorkSheet::from_worksheet(sheet_id, &new_name, target_id, copy_worksheet);
        self.sheets.push(worksheet);
        self.get_worksheet_mut(sheet_id)
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

    pub fn define_name(&mut self, name: &str, value: &str) -> WorkbookResult<()> {
        self.workbook.borrow_mut().defined_names.add_define_name(name, value, None);
        Ok(())
    }

    pub fn define_local_name(&mut self, name: &str, value: &str, sheet_id: u32) -> WorkbookResult<()> {
        if sheet_id > self.sheets.len() as u32 {
            return Err(WorkbookError::SheetError(WorkSheetError::FileNotFound));
        }
        self.workbook.borrow_mut().defined_names.add_define_name(name, value, Some(sheet_id - 1));
        Ok(())
    }

    /// Get the Range Reference for the given Workbook-level Name (if found)
    pub fn get_defined_name(&self, name: &str) -> WorkbookResult<String> {
        let ret = self.workbook.borrow();
        ret.defined_names
            .get_defined_name(name, None)
            .map(String::from) 
            .ok_or(WorkbookError::SheetError(WorkSheetError::FileNotFound))
    }
    /// Get the Range Reference for the given Worksheet-level Name (if found)
    pub fn get_defined_local_name(&self, name: &str, sheet_id: u32) -> WorkbookResult<String> {
        if sheet_id > self.sheets.len() as u32 {
            return Err(WorkbookError::SheetError(WorkSheetError::FileNotFound));
        }
        let ret = self.workbook.borrow();
        ret.defined_names
            .get_defined_name(name, Some(sheet_id - 1))
            .map(String::from) 
            .ok_or(WorkbookError::SheetError(WorkSheetError::FileNotFound))   
    }

    pub fn worksheets_mut(&mut self) -> slice::IterMut<WorkSheet> {
        self.sheets.iter_mut()
    }

    pub fn worksheets(&self) -> slice::Iter<WorkSheet> {
        self.sheets.iter()
    }

    pub fn read_only_recommended(&mut self) -> WorkbookResult<()> {
        let workbook = &mut self.workbook.borrow_mut();
        let mut file_sharing = workbook.file_sharing.take().unwrap_or_default();
        file_sharing.read_only_recommended = 1;
        workbook.file_sharing = Some(file_sharing);
        Ok(())
    }

    pub fn set_properties(&mut self, properties: &Properties) -> WorkbookResult<()> {
        let core_properties = self.get_core_properties();
        core_properties.update_by_properties(properties);
        let app_properties = self.get_app_properties();
        app_properties.update_by_properties(properties);
        Ok(())
    }
}

impl Workbook {
    fn from_path_v2<P: AsRef<Path>>(file_path: P) -> WorkbookResult<Workbook> {
        let file_name = file_path.as_ref().file_name().ok_or(ZipError::FileNotFound)?;
        let tmp_path = format!("./~${}_{:X}", file_name.to_str().ok_or(ZipError::FileNotFound)?, id_util::new_id());
        let file = File::open(&file_path)?;
        let mut archive = zip::ZipArchive::new(file)?;
        let mut medias = Medias::default();
        let workbook_xml = xml::workbook::Workbook::from_zip_file(&mut archive, "xl/workbook.xml");
        let workbook_rel = Relationships::from_zip_file(&mut archive, "xl/_rels/workbook.xml.rels");
        let content_types = ContentTypes::from_zip_file(&mut archive, "[Content_Types].xml");
        let style_sheet = StyleSheet::from_zip_file(&mut archive, "xl/styles.xml");
        let metadata = Metadata::from_zip_file(&mut archive, "xl/metadata.xml");
        let shared_string = SharedString::from_zip_file(&mut archive, "xl/sharedStrings.xml");
        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            let file_name = file.name();
            if file_name.starts_with("xl/media/") {
                medias.add_existed_media(&file_name);
            }
        }
        let workbook = Rc::new(RefCell::new(workbook_xml.unwrap_or_default()));
        let workbook_rel = Rc::new(RefCell::new(workbook_rel.unwrap_or_default()));
        let content_types = Rc::new(RefCell::new(content_types.unwrap_or_default()));
        let style_sheet = Rc::new(RefCell::new(style_sheet.unwrap_or_default()));
        let metadata = Rc::new(RefCell::new(metadata.unwrap_or_default()));
        let shared_string = Rc::new(shared_string.unwrap_or_default());
        let medias = Rc::new(RefCell::new(medias));
        let sheets = workbook.borrow().sheets.sheets.iter().map(
            |sheet_xml| {
                let binding = workbook_rel.borrow();
                let (target, target_id) = binding.get_target(&sheet_xml.r_id);
                WorkSheet::from_archive(
                    sheet_xml.sheet_id,
                    &sheet_xml.name,
                    target,
                    target_id,
                    &file_path,
                    &mut archive,
                    Rc::clone(&workbook),
                    Rc::clone(&workbook_rel),
                    Rc::clone(&style_sheet),
                    Rc::clone(&content_types),
                    Rc::clone(&medias),
                    Rc::clone(&metadata),
                    Rc::clone(&shared_string),
                )
            }).collect::<Vec<WorkSheet>>();
        let api_workbook = Workbook {
            sheets,
            tmp_path,
            file_path: file_path.as_ref().to_str().unwrap().to_string(),
            closed: false,
            workbook: Rc::clone(&workbook),
            workbook_rel: Rc::clone(&workbook_rel),
            style_sheet: Rc::clone(&style_sheet),
            content_types: Rc::clone(&content_types),
            medias: Rc::clone(&medias),
            metadata,
            core_properties: None,
            app_properties: None,
            shared_string,
        };
        Ok(api_workbook)
    }

    pub fn from_path<P: AsRef<Path>>(file_path: P) -> WorkbookResult<Workbook> {
        let workbook = Self::from_path_v2(file_path);
        // let workbook = block_on(workbook);
        workbook
    }

    async fn save_async(&self) -> WorkbookResult<()> {
        let workbook = self.workbook.borrow();
        let workbook = workbook.save_async(&self.tmp_path);
        let style_sheet = self.style_sheet.borrow();
        let style_sheet = style_sheet.save_async(&self.tmp_path);
        let workbook_rel = self.workbook_rel.borrow();
        let workbook_rel = workbook_rel.save_async(&self.tmp_path, XlsxFileType::WorkbookRels);
        let content_types = self.content_types.borrow();
        let content_types = content_types.save_async(&self.tmp_path);
        let medias = self.medias.borrow();
        let medias = medias.save_async(&self.tmp_path);
        let metadata = self.metadata.borrow();
        let metadata = metadata.save_async(&self.tmp_path);
        join!(workbook, style_sheet, workbook_rel, content_types, medias, metadata);
        Ok(())
    }

    pub fn save_as<P: AsRef<Path>>(&self, file_path: P) -> WorkbookResult<()> {
        if self.closed {
            return Err(WorkbookError::FileNotFound);
        }
        // Extract xlsx to tmp dir
        zip_util::extract_dir(&self.file_path, &self.tmp_path)?;
        // save sheets
        self.sheets.iter().for_each(|s| s.save_as(&self.tmp_path).unwrap());
        block_on(self.save_async()).unwrap();
        // save if modified
        if let Some(core_propertises) = &self.core_properties {
            core_propertises.save(&self.tmp_path);
        }
        if let Some(app_properties) = &self.app_properties {
            app_properties.save(&self.tmp_path);
        }
        // package files
        zip_util::zip_dir(&self.tmp_path, file_path)?;
        // clean cache
        fs::remove_dir_all(&self.tmp_path).unwrap();
        Ok(())
    }

    pub fn save(&mut self) -> WorkbookResult<()> {
        self.save_as(&self.file_path.clone())
    }

    pub fn finish(&mut self) {
        if !self.closed {
            fs::remove_dir_all(&self.tmp_path);
            self.closed = true;
        }
    }
}

impl Drop for Workbook {
    fn drop(&mut self) {
        self.finish();
    }
}