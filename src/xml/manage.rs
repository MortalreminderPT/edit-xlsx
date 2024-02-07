use std::collections::HashMap;
use std::path::Path;
use crate::result::{CellResult, RowResult};
use crate::xml::shared_string::SharedString;
use crate::xml::sheet_data::{Cell, Row};
use crate::xml::style::StyleSheet;
use crate::xml::workbook::{Sheet, Workbook};
use crate::xml::workbook_rel::Relationships;
use crate::xml::worksheet::WorkSheet;

#[derive(Debug)]
pub(crate) struct XmlManager {
    workbook: Workbook,
    workbook_rel: Relationships,
    worksheets: HashMap<u32, WorkSheet>,
    shared_string: SharedString,
    style_sheet: StyleSheet,
}

pub(crate) trait XmlIo<T> {
    fn from_path<P: AsRef<Path>>(file_path: P) -> T;
    fn save<P: AsRef<Path>>(&mut self, file_path: P);
}

pub(crate) trait Create {
    fn create_worksheet(&mut self) -> (u32, &mut WorkSheet);
}

pub(crate) trait Borrow {
    fn borrow_workbook(&self) -> &Workbook;
    fn borrow_worksheets(&self) -> &HashMap<u32, WorkSheet>;
    fn borrow_worksheet(&self, id: u32) -> &WorkSheet;
    fn borrow_shared_string(&self) -> &SharedString;

    fn borrow_style_sheet(&self) -> &StyleSheet;

    fn borrow_workbook_mut(&mut self) -> &mut Workbook;
    fn borrow_worksheets_mut(&mut self) -> &mut HashMap<u32, WorkSheet>;
    fn borrow_worksheet_mut(&mut self, id: u32) -> &mut WorkSheet;
    fn borrow_shared_string_mut(&mut self) -> &mut SharedString;
    fn borrow_style_sheet_mut(&mut self) -> &mut StyleSheet;
}

impl XmlIo<XmlManager> for XmlManager {
     fn from_path<P: AsRef<Path>>(path: P) -> XmlManager {
         let workbook = Workbook::from_path(&path);
         let workbook_rel = Relationships::from_path(&path);
         let shared_string = SharedString::from_path(&path);
         let style_sheet = StyleSheet::from_path(&path);
         let worksheets: HashMap<u32, WorkSheet> = workbook.sheets.sheets.iter()
             .map(|sheet| (sheet.sheet_id, WorkSheet::from_path(&path, sheet.sheet_id))).collect();
         XmlManager {
             workbook,
             workbook_rel,
             worksheets,
             shared_string,
             style_sheet,
         }
     }
    fn save<P: AsRef<Path>>(&mut self, file_path: P) {
        self.workbook.save(&file_path);
        self.worksheets.iter_mut().for_each(|(id, worksheet)| worksheet.save(&file_path, *id));
        self.shared_string.save(&file_path);
        self.style_sheet.save(&file_path);
        self.workbook_rel.save(&file_path);
    }
}

impl Create for XmlManager {
    fn create_worksheet(&mut self) -> (u32, &mut WorkSheet) {
        let id = self.workbook_rel.add_worksheet();
        let work_sheet = WorkSheet::new();
        self.worksheets.insert(id, work_sheet);
        self.workbook.sheets.sheets.push(
            Sheet::new(id)
        );
        (id, self.worksheets.get_mut(&id).unwrap())
    }
}

impl Borrow for XmlManager {
    fn borrow_workbook(&self) -> &Workbook {
        &self.workbook
    }
    fn borrow_worksheets(&self) -> &HashMap<u32, WorkSheet> {
        &self.worksheets
    }
    fn borrow_worksheet(&self, id: u32) -> &WorkSheet {
        self.worksheets.get(&id).unwrap()
    }
    fn borrow_shared_string(&self) -> &SharedString {
        &self.shared_string
    }

    fn borrow_style_sheet(&self) -> &StyleSheet {
        &self.style_sheet
    }

    fn borrow_workbook_mut(&mut self) -> &mut Workbook {
        &mut self.workbook
    }
    fn borrow_worksheets_mut(&mut self) -> &mut HashMap<u32, WorkSheet> {
        &mut self.worksheets
    }
    fn borrow_worksheet_mut(&mut self, id: u32) -> &mut WorkSheet {
        self.worksheets.get_mut(&id).unwrap()
    }

    fn borrow_shared_string_mut(&mut self) -> &mut SharedString {
        &mut self.shared_string
    }

    fn borrow_style_sheet_mut(&mut self) -> &mut StyleSheet {
        &mut self.style_sheet
    }
}

trait EditSheet {
    
}

pub(crate) trait EditRow {
    fn get(&mut self, row_id: u32) -> RowResult<&mut Row>;
    fn create(&mut self, row_id: u32) -> RowResult<&mut Row>;
    fn update(&mut self, row_id: u32) -> RowResult<&mut Row>;
    fn delete(&mut self, row_id: u32) -> RowResult<()>;
    fn sort(&mut self);
}

pub(crate) trait EditCell {
    fn get(&mut self, col_id: u32) -> CellResult<&mut Cell>;
    fn create(&mut self, col_id: u32) -> CellResult<&mut Cell>;
    fn update(&mut self, col_id: u32) -> CellResult<&mut Cell>;
    fn delete(&mut self, col_id: u32) -> CellResult<&mut Cell>;
}