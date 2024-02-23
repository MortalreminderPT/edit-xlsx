use crate::api::cell::location::LocationRange;
use crate::api::worksheet::format::_Format;
use crate::api::worksheet::WorkSheet;
use crate::Format;
use crate::result::WorkSheetResult;

pub trait Col: _Col {
    fn set_column<R: LocationRange>(&mut self, col_range: R, width: f64) -> WorkSheetResult<()> {
        self.set_column_all(col_range, width, None)
    }
    fn set_column_pixels<R: LocationRange>(&mut self, col_range: R, width: f64) -> WorkSheetResult<()> {
        self.set_column_all(col_range, 0.5 * width, None)
    }

    fn set_column_with_format<R: LocationRange>(&mut self, col_range: R, width: f64, format: &Format) -> WorkSheetResult<()> {
        self.set_column_all(col_range, width, Some(format))
    }
    fn set_column_pixels_with_format<R: LocationRange>(&mut self, col_range: R, width: f64, format: &Format) -> WorkSheetResult<()> {
        self.set_column_all(col_range, 0.5 * width, Some(format))
    }

    fn hide_column<R: LocationRange>(&mut self, col_range: R) -> WorkSheetResult<()> {
        self.hide_col(col_range)
    }
}

pub(crate) trait _Col {
    fn set_column_all<R: LocationRange>(&mut self, col_range: R, width: f64, format: Option<&Format>) -> WorkSheetResult<()>;

    fn hide_col<R: LocationRange>(&mut self, col_range: R) -> WorkSheetResult<()>;
}

impl _Col for WorkSheet {
    fn set_column_all<R: LocationRange>(&mut self, col_range: R, width: f64, format: Option<&Format>) -> WorkSheetResult<()> {
        let mut style = None;
        if let Some(format) = format {
            style = Some(self.add_format(format));
        }
        let worksheet = &mut self.worksheet;
        let (first_col, last_col) = col_range.to_col_range();
        worksheet.create_col(first_col, last_col, Some(width), style, None)?;
        Ok(())
    }

    fn hide_col<R: LocationRange>(&mut self, col_range: R) -> WorkSheetResult<()> {
        let (first_col, last_col) = col_range.to_col_range();
        self.worksheet.hide_col(first_col, last_col, Some(1))?;
        Ok(())
    }
}