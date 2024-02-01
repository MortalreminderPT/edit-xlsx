use std::collections::HashMap;
use std::path::Path;
use crate::result::{CellResult, RowResult};
use crate::xml::sheet_data::{Cell, Row};
use crate::xml::shared_string::SharedString;
use crate::xml::workbook::Workbook;
use crate::xml::worksheet::WorkSheet;

#[derive(Debug)]
pub(crate) struct XmlManager {
    workbook: Workbook,
    worksheets: HashMap<u32, WorkSheet>,
    shared_string: SharedString,
}

pub(crate) trait XmlIo<T> {
    fn from_path<P: AsRef<Path>>(file_path: P) -> T;
    fn save<P: AsRef<Path>>(&mut self, file_path: P);
}

pub(crate) trait Borrow {
    fn borrow_workbook(&self) -> &Workbook;
    fn borrow_worksheets(&self) -> &HashMap<u32, WorkSheet>;
    fn borrow_worksheet(&self, id: u32) -> &WorkSheet;
    fn borrow_shared_string(&self) -> &SharedString;
    fn borrow_workbook_mut(&mut self) -> &mut Workbook;
    fn borrow_worksheets_mut(&mut self) -> &mut HashMap<u32, WorkSheet>;
    fn borrow_worksheet_mut(&mut self, id: u32) -> &mut WorkSheet;
    fn borrow_shared_string_mut(&mut self) -> &mut SharedString;
}

impl XmlIo<XmlManager> for XmlManager {
     fn from_path<P: AsRef<Path>>(path: P) -> XmlManager {
        let workbook = Workbook::from_path(&path);
        let shared_string = SharedString::from_path(&path);
        let worksheets: HashMap<u32, WorkSheet> = workbook.sheets.sheets.iter()
            .map(|sheet| (sheet.sheet_id, WorkSheet::from_path(&path, sheet.sheet_id))).collect();
        XmlManager {
            workbook,
            worksheets,
            shared_string,
        }
    }
    fn save<P: AsRef<Path>>(&mut self, file_path: P) {
        self.workbook.save(&file_path);
        self.worksheets.iter_mut().for_each(|(id, worksheet)| worksheet.save(&file_path, *id));
        self.shared_string.save(&file_path);
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

trait EditCell {
    fn get(&mut self, col_id: u32) -> CellResult<&mut Cell>;
    fn create(&mut self, col_id: u32) -> CellResult<&mut Cell>;
    fn update(&mut self, col_id: u32) -> CellResult<&mut Cell>;
    fn delete(&mut self, col_id: u32) -> CellResult<&mut Cell>;
}