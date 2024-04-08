use std::{fs, slice};
use std::cell::RefCell;
use std::path::Path;
use std::rc::Rc;
use futures::executor::block_on;
use futures::join;
use crate::api::worksheet::WorkSheet;
use crate::file::XlsxFileType;
use crate::utils::zip_util;
use crate::result::{WorkSheetError, WorkbookError, WorkbookResult};
use crate::{Properties, xml};
use crate::xml::content_types::ContentTypes;
use crate::xml::core_properties::CoreProperties;
use crate::xml::app_properties::AppProperties;
use crate::xml::io::Io;
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
    workbook: Rc<RefCell<xml::workbook::Workbook>>,
    style_sheet: Rc<RefCell<StyleSheet>>,
    workbook_rel: Rc<RefCell<Relationships>>,
    content_types: Rc<RefCell<ContentTypes>>,
    medias: Rc<RefCell<Medias>>,
    metadata: Rc<RefCell<Metadata>>,
    core_properties: Option<CoreProperties>,
    app_properties: Option<AppProperties>,
    shared_string: Rc<SharedString>,
}

///
/// Private methods
///
impl Workbook {
    fn get_core_properties(&mut self) -> &mut CoreProperties {
        self.core_properties.get_or_insert(CoreProperties::from_path(&self.tmp_path).unwrap())
    }

    fn get_app_properties(&mut self) -> &mut AppProperties {
        self.app_properties.get_or_insert(AppProperties::from_path(&self.tmp_path).unwrap())
    }
}

impl Workbook {
    pub fn new() -> Workbook {
        let file_path = concat!(env!("CARGO_MANIFEST_DIR"), "/resources/new.xlsx");
        let mut wb = Self::from_path(file_path).unwrap();
        wb.file_path = String::from(file_path);
        wb
    }

    pub fn get_worksheet(&mut self, id: u32) -> WorkbookResult<&mut WorkSheet> {
        let sheet = self.sheets
            .iter_mut()
            .find(|sheet| sheet.id == id).ok_or(WorkSheetError::FileNotFound)?;
        Ok(sheet)
    }

    pub fn read_worksheet(&self, id: u32) -> WorkbookResult<&WorkSheet> {
        let sheet = self.sheets
            .iter()
            .find(|sheet| sheet.id == id).ok_or(WorkSheetError::FileNotFound)?;
        Ok(sheet)
    }

    pub fn add_worksheet(&mut self) -> WorkbookResult<& mut WorkSheet> {
        let sheet_id = self.workbook.borrow().next_sheet_id();
        let (r_id, target) = self.workbook_rel.borrow_mut().add_worksheet(sheet_id);
        // let name = self.names.next_sheet_name();
        let name = self.workbook.borrow_mut().add_worksheet(sheet_id, r_id)?;
        let sheet = WorkSheet::from_xml(
            sheet_id,
            &name,
            &target,
            &self.tmp_path,
            Rc::clone(&self.workbook),
            Rc::clone(&self.workbook_rel),
            // Weak::clone(&self.workbook_api),
            Rc::clone(&self.style_sheet),
            Rc::clone(&self.content_types),
            Rc::clone(&self.medias),
            Rc::clone(&self.metadata),
            Rc::clone(&self.shared_string),
        );
        self.sheets.push(sheet);
        self.get_worksheet(sheet_id)
    }

    pub fn add_worksheet_by_name(&mut self, name: &str) -> WorkbookResult<& mut WorkSheet> {
        let id = self.workbook.borrow().next_sheet_id();
        let (r_id, target) = self.workbook_rel.borrow_mut().add_worksheet(id);
        self.workbook.borrow_mut().add_worksheet_by_name(id, r_id, name)?;
        let sheet = WorkSheet::from_xml(
            id,
            name,
            &target,
            &self.tmp_path,
            Rc::clone(&self.workbook),
            Rc::clone(&self.workbook_rel),
            // Weak::clone(&self.workbook_api),
            Rc::clone(&self.style_sheet),
            Rc::clone(&self.content_types),
            Rc::clone(&self.medias),
            Rc::clone(&self.metadata),
            Rc::clone(&self.shared_string),
        );
        self.sheets.push(sheet);
        self.get_worksheet(id)
    }

    pub fn duplicate_worksheet(&mut self, id: u32) -> WorkbookResult<& mut WorkSheet> {
        let copy_worksheet = self.sheets
            .iter()
            .find(|sheet| sheet.id == id).ok_or(WorkSheetError::FileNotFound)?;
        let sheet_id = self.workbook.borrow().next_sheet_id();
        let (r_id, target) = self.workbook_rel.borrow_mut().add_worksheet(sheet_id);
        let name = self.workbook.borrow_mut().add_worksheet(sheet_id, r_id)?;
        let sheet = WorkSheet::from_worksheet(
            sheet_id,
            &name,
            &target,
            &self.tmp_path,
            Rc::clone(&self.workbook),
            Rc::clone(&self.workbook_rel),
            Rc::clone(&self.style_sheet),
            Rc::clone(&self.content_types),
            Rc::clone(&self.medias),
            Rc::clone(&self.metadata),
            copy_worksheet,
            Rc::clone(&self.shared_string),
        );
        self.sheets.push(sheet);
        self.get_worksheet(sheet_id)
    }

    pub fn duplicate_worksheet_by_name(&mut self, name: &str) -> WorkbookResult<& mut WorkSheet> {
        let copy_worksheet = self.sheets
            .iter()
            .find(|sheet| sheet.name == name).ok_or(WorkSheetError::FileNotFound)?;
        let sheet_id = self.workbook.borrow().next_sheet_id();
        let (r_id, target) = self.workbook_rel.borrow_mut().add_worksheet(sheet_id);
        let name = format!("{name} Duplicated");
        self.workbook.borrow_mut().add_worksheet_by_name(sheet_id, r_id, &name)?;
        let sheet = WorkSheet::from_worksheet(
            sheet_id,
            &name,
            &target,
            &self.tmp_path,
            Rc::clone(&self.workbook),
            Rc::clone(&self.workbook_rel),
            Rc::clone(&self.style_sheet),
            Rc::clone(&self.content_types),
            Rc::clone(&self.medias),
            Rc::clone(&self.metadata),
            copy_worksheet,
            Rc::clone(&self.shared_string),
        );
        self.sheets.push(sheet);
        self.get_worksheet(sheet_id)
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

    pub fn worksheets(&mut self) -> slice::IterMut<WorkSheet> {
        self.sheets.iter_mut()
    }

    pub fn get_worksheet_by_name(&mut self, name: &str) -> WorkbookResult<& mut WorkSheet> {
        let sheet = self.sheets
            .iter_mut()
            .find(|sheet| sheet.name == name);
        match sheet {
            Some(sheet) => Ok(sheet),
            None => Err(WorkbookError::SheetError(WorkSheetError::FileNotFound))
        }
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
    async fn from_path_async<P: AsRef<Path>>(file_path: P) -> WorkbookResult<Workbook> {
        let tmp_path = Workbook::extract_tmp_dir(&file_path)?;
        let workbook = xml::workbook::Workbook::from_path_async(&tmp_path);
        let workbook_rel = Relationships::from_path_async(&tmp_path, XlsxFileType::WorkbookRels);
        let style_sheet = StyleSheet::from_path_async(&tmp_path);
        let content_types = ContentTypes::from_path_async(&tmp_path);
        let medias = Medias::from_path_async(&tmp_path);
        let metadata = Metadata::from_path_async(&tmp_path);
        let shared_string = SharedString::from_path_async(&tmp_path);
        let (workbook, workbook_rel, style_sheet,
            content_types, medias, metadata,
            shared_string
        ) = join!(workbook, workbook_rel, style_sheet, content_types, medias, metadata, shared_string);

        let workbook = Rc::new(RefCell::new(workbook.unwrap()));
        let workbook_rel = Rc::new(RefCell::new(workbook_rel.unwrap()));
        let style_sheet = Rc::new(RefCell::new(style_sheet.unwrap()));
        let content_types = Rc::new(RefCell::new(content_types.unwrap()));
        let medias = Rc::new(RefCell::new(medias.unwrap()));
        let metadata = Rc::new(RefCell::new(metadata.unwrap_or_default()));
        let shared_string = Rc::new(shared_string.unwrap_or_default());

        let sheets = workbook.borrow().sheets.sheets.iter().map(
            |sheet_xml| {
                WorkSheet::from_xml(
                    sheet_xml.sheet_id,
                    &sheet_xml.name,
                    &workbook_rel.borrow().get_target(&sheet_xml.r_id),
                    &tmp_path,
                    Rc::clone(&workbook),
                    Rc::clone(&workbook_rel),
                    Rc::clone(&style_sheet),
                    Rc::clone(&content_types),
                    Rc::clone(&medias),
                    Rc::clone(&metadata),
                    Rc::clone(&shared_string),
                )
            }).collect::<Vec<WorkSheet>>();

        let workbook = Workbook {
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

        Ok(workbook)
    }

    pub fn from_path<P: AsRef<Path>>(file_path: P) -> WorkbookResult<Workbook> {
        let workbook = Self::from_path_async(file_path);
        let workbook = block_on(workbook);
        workbook
    }

    fn extract_tmp_dir<P: AsRef<Path>>(file_path: P) -> WorkbookResult<String> {
        Ok(zip_util::extract_dir(file_path)?)
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