use std::{fs, slice, thread};
use std::cell::RefCell;
use std::path::Path;
use std::rc::Rc;
use std::sync::Arc;
use std::thread::JoinHandle;
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
}

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
        let mut wb = Self::from_path("resource/new.xlsx").unwrap();
        wb.file_path = String::from("./new.xlsx");
        wb
    }

    pub fn get_worksheet(&mut self, id: u32) -> WorkbookResult<& mut WorkSheet> {
        let sheet = self.sheets
            .iter_mut()
            .find(|sheet| sheet.id == id).ok_or(WorkSheetError::FileNotFound)?;
        Ok(sheet)
    }

    pub fn add_worksheet(&mut self) -> WorkbookResult<& mut WorkSheet> {
        let id = self.workbook.borrow().next_sheet_id();
        let (r_id, target) = self.workbook_rel.borrow_mut().add_worksheet(id);
        // let name = self.names.next_sheet_name();
        let name = self.workbook.borrow_mut().add_worksheet(id, r_id)?;
        let sheet = WorkSheet::from_xml(
            id,
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
        );
        self.sheets.push(sheet);
        self.get_worksheet(id)
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

#[derive(Debug)]
struct XmlHandler {
    workbook: JoinHandle<xml::workbook::Workbook>,
    workbook_rel: JoinHandle<Relationships>,
    style_sheet: JoinHandle<StyleSheet>,
    content_types: JoinHandle<ContentTypes>,
    medias: JoinHandle<Medias>,
    metadata: JoinHandle<Metadata>,
}

impl XmlHandler {
    fn from_path<P: AsRef<Path>>(file_path: P) -> WorkbookResult<XmlHandler> {
        let tmp_path = Workbook::extract_tmp_dir(&file_path)?;
        let tmp_path = Arc::new(tmp_path);
        Ok(XmlHandler {
            workbook: {
                let tmp_path = Arc::clone(&tmp_path);
                thread::spawn(move || {
                    xml::workbook::Workbook::from_path(&tmp_path.as_ref()).unwrap()
                })
            },
            workbook_rel: {
                let tmp_path = Arc::clone(&tmp_path);
                thread::spawn(move || {
                    Relationships::from_path(&tmp_path.as_ref(), XlsxFileType::WorkbookRels).unwrap()
                })
            },
            style_sheet: {
                let tmp_path = Arc::clone(&tmp_path);
                thread::spawn(move || {
                    StyleSheet::from_path(&tmp_path.as_ref()).unwrap()
                })
            },
            content_types: {
                let tmp_path = Arc::clone(&tmp_path);
                thread::spawn(move || {
                    ContentTypes::from_path(&tmp_path.as_ref()).unwrap()
                })
            },
            medias: {
                let tmp_path = Arc::clone(&tmp_path);
                thread::spawn(move || {
                    Medias::from_path(&tmp_path.as_ref()).unwrap()
                })
            },
            metadata: {
                let tmp_path = Arc::clone(&tmp_path);
                thread::spawn(move || {
                    Metadata::from_path(&tmp_path.as_ref()).unwrap_or_default()
                })
            },
        })
    }

    fn save<P: AsRef<Path>>(
        tmp_path: P,
        workbook: &'static xml::workbook::Workbook,
        workbook_rel: Relationships,
        style_sheet: StyleSheet,
        content_types: ContentTypes,
        medias: Medias,
        metadata: Metadata,
    ) -> WorkbookResult<()> {
        let tmp_path = Arc::new(tmp_path.as_ref().to_str().unwrap().to_string());
        let mut handle = vec![];
        {
            let tmp_path = Arc::clone(&tmp_path);
            handle.push(thread::spawn(move || {
                workbook.save(tmp_path.as_ref());
            }));
        }
        {
            let tmp_path = Arc::clone(&tmp_path);
            handle.push(thread::spawn(move || {
                style_sheet.save(&tmp_path.as_ref());
            }));
        }
        {
            let tmp_path = Arc::clone(&tmp_path);
            handle.push(thread::spawn(move || {
                workbook_rel.save(tmp_path.as_ref(), XlsxFileType::WorkbookRels);
            }));
        }
        {
            let tmp_path = Arc::clone(&tmp_path);
            handle.push(thread::spawn(move || {
                content_types.save(tmp_path.as_ref());
            }));
        }
        {
            let tmp_path = Arc::clone(&tmp_path);
            handle.push(thread::spawn(move || {
                medias.save(tmp_path.as_ref());
            }));
        }
        {
            let tmp_path = Arc::clone(&tmp_path);
            handle.push(thread::spawn(move || {
                metadata.save(tmp_path.as_ref());
            }));
        }
        for h in handle {
            h.join().unwrap();
        }
        Ok(())
    }
}

impl Workbook {
    pub fn from_path<P: AsRef<Path>>(file_path: P) -> WorkbookResult<Workbook> {
        let tmp_path = Workbook::extract_tmp_dir(&file_path)?;
        // get xmls from path
        let xml_handler = XmlHandler::from_path(&file_path)?;
        let workbook = Rc::new(RefCell::new(xml_handler.workbook.join().unwrap()));
        let workbook_rel = Rc::new(RefCell::new(xml_handler.workbook_rel.join().unwrap()));
        let style_sheet = Rc::new(RefCell::new(xml_handler.style_sheet.join().unwrap()));
        let content_types = Rc::new(RefCell::new(xml_handler.content_types.join().unwrap()));
        let medias = Rc::new(RefCell::new(xml_handler.medias.join().unwrap()));
        let metadata = Rc::new(RefCell::new(xml_handler.metadata.join().unwrap()));
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
                )
            }).collect::<Vec<WorkSheet>>();
        let mut workbook = Workbook {
            sheets,
            tmp_path,
            file_path: file_path.as_ref().to_str().unwrap().to_string(),
            closed: false,
            // workbook_api: Weak::new(),
            workbook: Rc::clone(&workbook),
            workbook_rel: Rc::clone(&workbook_rel),
            style_sheet: Rc::clone(&style_sheet),
            content_types: Rc::clone(&content_types),
            medias: Rc::clone(&medias),
            metadata,
            core_properties: None,
            app_properties: None,
        };
        // workbook.workbook_api = Rc::downgrade(&Rc::new(workbook));
        // Rc::downgrade(&Rc::new(RefCell::new(workbook)));
        Ok(workbook)
    }

    fn extract_tmp_dir<P: AsRef<Path>>(file_path: P) -> WorkbookResult<String> {
        Ok(zip_util::extract_dir(file_path)?)
    }

    pub fn save_as<P: AsRef<Path>>(&self, file_path: P) -> WorkbookResult<()> {
        // save sheets
        self.sheets.iter().for_each(|s| s.save_as(&self.tmp_path).unwrap());
        // save files
        // XmlHandler::save(&self.tmp_path,
        //                  self.workbook.take(), //.borrow().deref(),
        //                  self.workbook_rel.take(),
        //                  self.style_sheet.take(),
        //                  self.content_types.take(),
        //                  self.medias.take(),
        //                  self.metadata.take(),
        // )?;

        self.workbook.borrow().save(&self.tmp_path);
        self.style_sheet.borrow().save(&self.tmp_path);
        self.workbook_rel.borrow().save(&self.tmp_path, XlsxFileType::WorkbookRels);
        self.content_types.borrow().save(&self.tmp_path);
        self.medias.borrow().save(&self.tmp_path);
        self.metadata.borrow().save(&self.tmp_path);
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