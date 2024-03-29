use crate::api::cell::location::LocationRange;
use crate::api::worksheet::format::_Format;
use crate::api::worksheet::row::RowSet;
use crate::api::worksheet::WorkSheet;
use crate::Format;
use crate::result::WorkSheetResult;

pub trait Col: _Col {
    fn set_column<R: LocationRange>(&mut self, col_range: R, width: f64) -> WorkSheetResult<()> {
        let mut col_set = ColSet::default();
        col_set.width = Some(width);
        self.set_col_by_colset(col_range, &col_set)?;
        Ok(())
    }
    fn set_column_pixels<R: LocationRange>(&mut self, col_range: R, width: f64) -> WorkSheetResult<()> {
        let mut col_set = ColSet::default();
        col_set.width = Some(0.5 * width);
        self.set_col_by_colset(col_range, &col_set)?;
        Ok(())
    }

    fn set_column_with_format<R: LocationRange>(&mut self, col_range: R, width: f64, format: &Format) -> WorkSheetResult<()> {
        let mut col_set = ColSet::default();
        col_set.style = Some(self.add_format(format));
        col_set.width = Some(width);
        self.set_col_by_colset(col_range, &col_set)?;
        Ok(())
    }
    fn set_column_pixels_with_format<R: LocationRange>(&mut self, col_range: R, width: f64, format: &Format) -> WorkSheetResult<()> {
        let mut col_set = ColSet::default();
        col_set.style = Some(self.add_format(format));
        col_set.width = Some(0.5 * width);
        self.set_col_by_colset(col_range, &col_set)?;
        Ok(())
    }

    fn hide_column<R: LocationRange>(&mut self, col_range: R) -> WorkSheetResult<()> {
        let mut col_set = ColSet::default();
        col_set.hidden = Some(1);
        self.set_col_by_colset(col_range, &col_set)?;
        Ok(())
    }

    fn set_column_level<R: LocationRange>(&mut self, col_range: R, level: u32) -> WorkSheetResult<()> {
        let mut col_set = ColSet::default();
        col_set.outline_level = Some(level);
        self.set_col_by_colset(col_range, &col_set)?;
        Ok(())
    }

    fn collapse_col<R: LocationRange>(&mut self, col_range: R) -> WorkSheetResult<()> {
        let mut col_set = ColSet::default();
        col_set.collapsed = Some(1);
        self.set_col_by_colset(col_range, &col_set)?;
        Ok(())
    }
}

pub(crate) trait _Col: _Format {
    fn set_col_by_colset<R: LocationRange>(&mut self, col_range: R, col_set: &ColSet) -> WorkSheetResult<()>;
}

impl _Col for WorkSheet {
    fn set_col_by_colset<R: LocationRange>(&mut self, col_range: R, col_set: &ColSet) -> WorkSheetResult<()> {
        // let (start, end) = col_range.to_col_range();
        // for col in start..=end {
            self.worksheet.set_col_by_colset(col_range, col_set)?;
        // }
        Ok(())
    }
}

#[derive(Debug, Default)]
pub(crate) struct ColSet {
    pub(crate) width: Option<f64>,
    pub(crate) style: Option<u32>,
    pub(crate) outline_level: Option<u32>,
    pub(crate) hidden: Option<u8>,
    pub(crate) collapsed: Option<u8>,
}