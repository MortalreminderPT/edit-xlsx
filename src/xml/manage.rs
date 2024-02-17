use std::cell::{Ref, RefCell, RefMut};
use std::collections::HashMap;
use std::io;
use std::path::Path;
use std::rc::Rc;
use crate::result::{CellResult, SheetError, SheetResult};
use crate::xml::shared_string::SharedString;
use crate::xml::sheet_data::{Cell, Row};
use crate::xml::style::StyleSheet;
use crate::xml::workbook::{Sheet, Workbook};
use crate::xml::workbook_rel::Relationships;
use crate::xml::worksheet::WorkSheet;

#[derive(Debug, Default)]
pub(crate) struct XmlManager {
    pub(crate) workbook: Rc<RefCell<Workbook>>,
    pub(crate) workbook_rel: Rc<RefCell<Relationships>>,
    pub(crate) worksheets: Rc<RefCell<HashMap<u32, WorkSheet>>>,
    shared_string: SharedString,
    pub(crate) style_sheet: Rc<RefCell<StyleSheet>>,
}

pub(crate) trait XmlIo<T> {
    fn from_path<P: AsRef<Path>>(file_path: P) -> io::Result<T>;
    fn save<P: AsRef<Path>>(&mut self, file_path: P);
}

trait EditSheet {
    
}

pub(crate) trait EditCell {
    fn get(&mut self, col_id: u32) -> CellResult<&mut Cell>;
    fn create(&mut self, col_id: u32) -> CellResult<&mut Cell>;
    fn update(&mut self, col_id: u32) -> CellResult<&mut Cell>;
    fn delete(&mut self, col_id: u32) -> CellResult<&mut Cell>;
}