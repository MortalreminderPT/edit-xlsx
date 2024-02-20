use crate::api::worksheet::format::_Format;
use crate::api::worksheet::Sheet;
use crate::Format;
use crate::result::SheetResult;

pub trait Col: _Col {
    fn set_column(&mut self, first_col: u32, last_col: u32, width: f64) -> SheetResult<()> {
        self.set_column_all(first_col, last_col, width, None)
    }
    fn set_column_pixels(&mut self, first_col: u32, last_col: u32, width: f64) -> SheetResult<()> {
        self.set_column_all(first_col, last_col, 0.5 * width, None)
    }

    fn set_column_with_format(&mut self, first_col: u32, last_col: u32, width: f64, format: &Format) -> SheetResult<()> {
        self.set_column_all(first_col, last_col, width, Some(format))
    }
    fn set_column_pixels_with_format(&mut self, first_col: u32, last_col: u32, width: f64, format: &Format) -> SheetResult<()> {
        self.set_column_all(first_col, last_col, 0.5 * width, Some(format))
    }
}

pub(crate) trait _Col {
    fn set_column_all(&mut self, first_col: u32, last_col: u32, width: f64, format: Option<&Format>) -> SheetResult<()>;
}

impl _Col for Sheet {
    fn set_column_all(&mut self, first_col: u32, last_col: u32, width: f64, format: Option<&Format>) -> SheetResult<()> {
        let mut style = None;
        if let Some(format) = format {
            style = Some(self.add_format(format));
        }
        let worksheet = &mut self.worksheet;
        worksheet.create_col(first_col, last_col, Some(width), style, None)?;
        Ok(())
    }
}