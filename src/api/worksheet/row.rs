use crate::Format;
use crate::api::worksheet::format::_Format;
use crate::api::worksheet::WorkSheet;
use crate::result::WorkSheetResult;

pub trait Row: _Row {
    fn set_row(&mut self, row: u32, height: f64) -> WorkSheetResult<()> {
        let mut row_set = RowSet::default();
        row_set.height = Some(height);
        self.set_row_by_rowset(row, &row_set);
        Ok(())
    }

    fn set_row_pixels(&mut self, row: u32, height: f64) -> WorkSheetResult<()> {
        let mut row_set = RowSet::default();
        row_set.height = Some(0.5 * height);
        self.set_row_by_rowset(row, &row_set);
        Ok(())
    }
    fn set_row_with_format(&mut self, row: u32, height: f64, format: &Format) -> WorkSheetResult<()> {
        let mut row_set = RowSet::default();
        row_set.height = Some(height);
        row_set.style = Some(self.add_format(format));
        self.set_row_by_rowset(row, &row_set);
        Ok(())
    }

    fn set_row_pixels_with_format(&mut self, row: u32, height: f64, format: &Format) -> WorkSheetResult<()> {
        let mut row_set = RowSet::default();
        row_set.height = Some(0.5 * height);
        row_set.style = Some(self.add_format(format));
        self.set_row_by_rowset(row, &row_set);
        Ok(())
    }

    fn hide_row(&mut self, row: u32) -> WorkSheetResult<()> {
        let mut row_set = RowSet::default();
        row_set.hidden = Some(1);
        self.set_row_by_rowset(row, &row_set);
        Ok(())
    }

    fn set_row_level(&mut self, row: u32, level: u32) -> WorkSheetResult<()> {
        let mut row_set = RowSet::default();
        row_set.outline_level = Some(level);
        self.set_row_by_rowset(row, &row_set);
        Ok(())
    }

    fn collapse_row(&mut self, row: u32) -> WorkSheetResult<()> {
        let mut row_set = RowSet::default();
        row_set.collapsed = Some(1);
        self.set_row_by_rowset(row, &row_set);
        Ok(())
    }
}

pub(crate) trait _Row: _Format {
    fn set_row_by_rowset(&mut self, row: u32, row_set: &RowSet);
}

impl _Row for WorkSheet {
    fn set_row_by_rowset(&mut self, row: u32, row_set: &RowSet) {
        self.worksheet.sheet_data.set_row_by_rowset(row, row_set);
    }
}

///
/// Manages the Row settings that are modified,
/// only the parts that are not None are applied to the Row's modifications.
///
#[derive(Debug, Default)]
pub(crate) struct RowSet {
    pub(crate) height: Option<f64>,
    pub(crate) style: Option<u32>,
    pub(crate) outline_level: Option<u32>,
    pub(crate) hidden: Option<u8>,
    pub(crate) collapsed: Option<u8>,
}