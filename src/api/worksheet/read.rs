use crate::api::cell::location::Location;
use crate::{WorkSheet, WorkSheetResult};
use crate::api::cell::values::CellType;
use crate::result::{CellError, RowError, WorkSheetError};

pub trait Read: _Read {
    // fn read_text<L: Location>(&self, loc: L) -> WorkSheetResult<&String> {
    //     self.read_value(loc)
    // }
    fn read<L: Location>(&self, loc: L) -> WorkSheetResult<&str> { self.read_text(loc) }
}

trait _Read {
    fn read_value<L: Location>(&self, loc: L) -> WorkSheetResult<&String>;
    fn read_text<L: Location>(&self, loc: L) -> WorkSheetResult<&str>;
}

impl _Read for WorkSheet {
    fn read_value<L: Location>(&self, loc: L) -> WorkSheetResult<&String> {
        let worksheet = &self.worksheet;
        let sheet_data = &worksheet.sheet_data;
        let value = sheet_data.get_value(&loc);
        value.ok_or(WorkSheetError::RowError(RowError::CellError(CellError::CellNotFound)))
    }

    fn read_text<L: Location>(&self, loc: L) -> WorkSheetResult<&str> {
        let worksheet = &self.worksheet;
        let sheet_data = &worksheet.sheet_data;
        let value = sheet_data.get_value(&loc);
        let cell_type = sheet_data.get_cell_type(&loc);
        match cell_type {
            Some(CellType::SharedString) => {
                Ok("SharedString")
            }
            _ => {
                if let Some(value) = value {
                    Ok(value.as_str())
                }
                else {
                    Err(WorkSheetError::RowError(RowError::CellError(CellError::CellNotFound)))
                }
            }
        }
    }
}