use crate::api::cell::location::Location;
use crate::{Cell, Format, WorkSheet, WorkSheetResult};
use crate::api::cell::values::{CellDisplay, CellType, CellValue};
use crate::api::worksheet::format::_Format;
use crate::api::worksheet::hyperlink::_Hyperlink;
use crate::result::{CellError, RowError, WorkSheetError};

pub trait Read: _Read {
    fn read_cell<L: Location>(&self, loc: L) -> WorkSheetResult<Cell<String>> {
        self.read_api_cell(&loc)
    }
    // fn read_text<L: Location>(&self, loc: L) -> WorkSheetResult<&str> { self.read_value(loc) }
    // fn read_string<L: Location>(&self, loc: L) -> WorkSheetResult<&str> { self.read_value(loc) }
    // fn read_shared_string<L: Location>(&self, loc: L) -> WorkSheetResult<&str> { self.read_value(loc) }
    // fn read_url<L: Location>(&self, loc: L) -> WorkSheetResult<&str> { Ok("") }
    // fn read_format<L: Location>(&self, loc: L) -> WorkSheetResult<Format> {self.read_format_all(loc)}
}

trait _Read {
    fn read_api_cell<L: Location>(&self, loc: &L) -> WorkSheetResult<Cell<String>>;
    // fn get_cell_type<L: Location>(&self, loc: L) -> WorkSheetResult<&CellType>;
    // fn read_value<L: Location>(&self, loc: L) -> WorkSheetResult<&str>;
    // fn read_text<L: Location>(&self, loc: L) -> WorkSheetResult<&str>;
    // fn read_format_all<L: Location>(&self, loc: L) -> WorkSheetResult<Format>;
}

impl _Read for WorkSheet {
    fn read_api_cell<L: Location>(&self, loc: &L) -> WorkSheetResult<Cell<String>> {
        let mut cell = self.worksheet.sheet_data.read_api_cell(loc)?;
        if let Some(style) = cell.style {
            cell.format = Some(self.get_format(style));
        }
        if let Some(CellType::SharedString) = cell.cell_type {
            let id: usize = if let Some(s) = cell.text {
                s.parse().unwrap_or_default()
            } else { 0 };
            cell.cell_type = Some(CellType::String);
            cell.text = Some(self.shared_string.get_text(id).unwrap().to_string());
        };
        cell.hyperlink = self.worksheet.get_hyperlink(loc);
        Ok(cell)
    }

    // fn get_cell_type<L: Location>(&self, loc: L) -> WorkSheetResult<&CellType> {
    //     let worksheet = &self.worksheet;
    //     let sheet_data = &worksheet.sheet_data;
    //     let cell_type = sheet_data.get_cell_type(&loc);
    //     cell_type.ok_or(WorkSheetError::RowError(RowError::CellError(CellError::CellNotFound)))
    // }
    // fn read_value<L: Location>(&self, loc: L) -> WorkSheetResult<&str> {
    //     let worksheet = &self.worksheet;
    //     let sheet_data = &worksheet.sheet_data;
    //     let cell_type = sheet_data.get_cell_type(&loc);
    //     let value = sheet_data.get_value(&loc);
    //     let text = match cell_type {
    //         Some(CellType::SharedString) => {
    //             let id: usize = value.unwrap_or("0").parse().unwrap();
    //             self.shared_string.get_text(id)
    //         },
    //         _ => value
    //     };
    //     text.ok_or(WorkSheetError::RowError(RowError::CellError(CellError::CellNotFound)))
    // }
    // fn read_text<L: Location>(&self, loc: L) -> WorkSheetResult<&str> {
    //     let worksheet = &self.worksheet;
    //     let sheet_data = &worksheet.sheet_data;
    //     let value = sheet_data.get_value(&loc);
    //     let cell_type = sheet_data.get_cell_type(&loc);
    //     match cell_type {
    //         Some(CellType::SharedString) => {
    //             Ok("SharedString")
    //         }
    //         _ => {
    //             value.ok_or(WorkSheetError::RowError(RowError::CellError(CellError::CellNotFound)))
    //         }
    //     }
    // }
    // fn read_format_all<L: Location>(&self, loc: L) -> WorkSheetResult<Format> {
    //     let worksheet = &self.worksheet;
    //     // let sheet_data = &worksheet.sheet_data;
    //     match worksheet.get_default_style(&loc) {
    //         Some(style) => Ok(self.get_format(style)),
    //         None => Err(WorkSheetError::FileNotFound)
    //     }
    // }
}