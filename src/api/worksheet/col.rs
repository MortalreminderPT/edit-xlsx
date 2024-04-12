use std::cell::RefCell;
pub(crate) use crate::api::cell::location::LocationRange;
pub(crate) use crate::api::worksheet::column::Column;
use crate::api::worksheet::format::_Format;
use crate::api::worksheet::row::RowSet;
use crate::api::worksheet::WorkSheet;
use crate::Format;
use crate::result::WorkSheetResult;

pub trait Col: _Col {
    fn set_column<R: LocationRange>(&mut self, col_range: R, width: f64) -> WorkSheetResult<()> {
        let mut col_set = Column::default();
        col_set.width = Some(width);
        self.set_by_column(col_range, &col_set)?;
        Ok(())
    }

    fn get_columns<R: LocationRange>(&self, col_range: R) -> WorkSheetResult<Vec<(u32, u32, Column)>> {
        self.list_by_range(col_range)
    }

    fn get_columns_with_format<R: LocationRange>(&self, col_range: R) -> WorkSheetResult<Vec<(u32, u32, Column, Option<Format>)>> {
        let all = self.list_by_range(col_range)?
            .iter()
            .map(|(min, max, col)| (
                *min,
                *max,
                *col,
                match col.style {
                    None => None,
                    Some(style) => Some(self.get_format(style))
                }
            ))
            .collect();
        Ok(all)
    }

    fn set_columns_with_format<R: LocationRange>(&mut self, col_range: R, column: &mut Column, format: &Format) -> WorkSheetResult<()> {
        column.style = Some(self.add_format(format));
        self.set_by_column(col_range, &column)?;
        Ok(())
    }

    fn set_columns<R: LocationRange>(&mut self, col_range: R, column: &mut Column) -> WorkSheetResult<()> {
        self.set_by_column(col_range, &column)?;
        Ok(())
    }

    fn get_column_width<R: LocationRange>(&self, col_range: R) -> WorkSheetResult<Vec<(u32, u32, Option<f64>)>> {
        let columns = self.list_by_range(col_range)?;
        let widths = columns.iter().map(|(min, max, col)| (*min, *max, col.width)).collect();
        Ok(widths)
    }

    fn set_column_pixels<R: LocationRange>(&mut self, col_range: R, width: f64) -> WorkSheetResult<()> {
        let mut col_set = Column::default();
        col_set.width = Some(0.5 * width);
        self.set_by_column(col_range, &col_set)?;
        Ok(())
    }
    fn set_column_with_format<R: LocationRange>(&mut self, col_range: R, width: f64, format: &Format) -> WorkSheetResult<()> {
        let mut col_set = Column::default();
        col_set.style = Some(self.add_format(format));
        col_set.width = Some(width);
        self.set_by_column(col_range, &col_set)?;
        Ok(())
    }
    fn set_column_pixels_with_format<R: LocationRange>(&mut self, col_range: R, width: f64, format: &Format) -> WorkSheetResult<()> {
        let mut col_set = Column::default();
        col_set.style = Some(self.add_format(format));
        col_set.width = Some(0.5 * width);
        self.set_by_column(col_range, &col_set)?;
        Ok(())
    }
    fn hide_column<R: LocationRange>(&mut self, col_range: R) -> WorkSheetResult<()> {
        let mut col_set = Column::default();
        col_set.hidden = Some(1);
        self.set_by_column(col_range, &col_set)?;
        Ok(())
    }
    fn set_column_level<R: LocationRange>(&mut self, col_range: R, level: u32) -> WorkSheetResult<()> {
        let mut col_set = Column::default();
        col_set.outline_level = Some(level);
        self.set_by_column(col_range, &col_set)?;
        Ok(())
    }
    fn collapse_col<R: LocationRange>(&mut self, col_range: R) -> WorkSheetResult<()> {
        let mut col_set = Column::default();
        col_set.collapsed = Some(1);
        self.set_by_column(col_range, &col_set)?;
        Ok(())
    }
}

pub(crate) trait _Col: _Format {
    fn set_by_column<R: LocationRange>(&mut self, col_range: R, col_set: &Column) -> WorkSheetResult<()>;
    fn list_by_range<R: LocationRange>(&self, col_range: R) -> WorkSheetResult<Vec<(u32, u32, Column)>>;
}

impl _Col for WorkSheet {
    fn set_by_column<R: LocationRange>(&mut self, col_range: R, col_set: &Column) -> WorkSheetResult<()> {
        self.worksheet.set_col_by_column(col_range, col_set)?;
        Ok(())
    }

    fn list_by_range<R: LocationRange>(&self, col_range: R) -> WorkSheetResult<Vec<(u32, u32, Column)>> {
        let columns = self.worksheet.get_col(col_range)?;
        Ok(columns)
    }
}
