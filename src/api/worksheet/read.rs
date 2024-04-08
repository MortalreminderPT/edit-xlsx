use crate::api::cell::location::Location;
use crate::{WorkSheet, WorkSheetResult};
use crate::api::cell::values::CellType;
use crate::result::{CellError, RowError, WorkSheetError};

pub trait Read: _Read {
    fn read<L: Location>(&self, loc: L) -> WorkSheetResult<&str> { self.read_value(loc) }
    fn read_text<L: Location>(&self, loc: L) -> WorkSheetResult<&str> { self.read_value(loc) }
    fn read_string<L: Location>(&self, loc: L) -> WorkSheetResult<&str> { self.read_value(loc) }
    fn read_shared_string<L: Location>(&self, loc: L) -> WorkSheetResult<&str> { self.read_value(loc) }
    fn read_number<L: Location>(&self, loc: L) -> WorkSheetResult<i32> { Ok(0) }
    fn read_double<L: Location>(&self, loc: L) -> WorkSheetResult<f64> { Ok(0.0) }
    fn read_boolean<L: Location>(&self, loc: L) -> WorkSheetResult<bool> { Ok(false) }
    fn read_url<L: Location>(&self, loc: L) -> WorkSheetResult<&str> { Ok("") }
}

trait _Read {
    fn get_cell_type<L: Location>(&self, loc: L) -> WorkSheetResult<&CellType>;
    fn read_value<L: Location>(&self, loc: L) -> WorkSheetResult<&str>;
    fn read_text<L: Location>(&self, loc: L) -> WorkSheetResult<&str>;
}

impl _Read for WorkSheet {
    fn get_cell_type<L: Location>(&self, loc: L) -> WorkSheetResult<&CellType> {
        let worksheet = &self.worksheet;
        let sheet_data = &worksheet.sheet_data;
        let cell_type = sheet_data.get_cell_type(&loc);
        cell_type.ok_or(WorkSheetError::RowError(RowError::CellError(CellError::CellNotFound)))
    }

    fn read_value<L: Location>(&self, loc: L) -> WorkSheetResult<&str> {
        let worksheet = &self.worksheet;
        let sheet_data = &worksheet.sheet_data;
        let cell_type = sheet_data.get_cell_type(&loc);
        let value = sheet_data.get_value(&loc);
        let text = match cell_type {
            Some(CellType::SharedString) => {
                let id: usize = value.unwrap_or("0").parse().unwrap();
                self.shared_string.get_text(id)
            },
            _ => value
        };
        text.ok_or(WorkSheetError::RowError(RowError::CellError(CellError::CellNotFound)))
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
                value.ok_or(WorkSheetError::RowError(RowError::CellError(CellError::CellNotFound)))
            }
        }
    }
}