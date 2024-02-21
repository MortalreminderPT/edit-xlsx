use crate::Format;
use crate::api::worksheet::format::_Format;
use crate::api::worksheet::WorkSheet;
use crate::result::WorkSheetResult;

pub trait Row: _Row {
    fn set_row(&mut self, row: u32, height: f64) -> WorkSheetResult<()> {
        self.set_row_all(row, height, None)
    }
    fn set_row_pixels(&mut self, row: u32, height: f64) -> WorkSheetResult<()> {
        self.set_row_all(row, 0.5 * height, None)
    }

    fn set_row_with_format(&mut self, row: u32, height: f64, format: &Format) -> WorkSheetResult<()> {
        self.set_row_all(row, height, Some(format))
    }
    fn set_row_pixels_with_format(&mut self, row: u32, height: f64, format: &Format) -> WorkSheetResult<()> {
        self.set_row_all(row, 0.5 * height, Some(format))
    }
}

pub(crate) trait _Row: _Format {
    fn set_row_all(&mut self, row: u32, height: f64, format: Option<&Format>) -> WorkSheetResult<()>;
}

impl _Row for WorkSheet {
    fn set_row_all(&mut self, row: u32, height: f64, format: Option<&Format>) -> WorkSheetResult<()> {
        let mut style = None;
        if let Some(format) = format {
            style = Some(self.add_format(format));
        }
        let worksheet = &mut self.worksheet;
        let sheet_data = &mut worksheet.sheet_data;
        sheet_data.set_row(row, height, style)?;
        Ok(())
    }
}