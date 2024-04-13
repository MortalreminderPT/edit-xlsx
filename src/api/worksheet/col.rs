use std::collections::HashMap;
pub(crate) use crate::api::cell::location::LocationRange;
pub(crate) use crate::api::worksheet::column::Column;
use crate::api::worksheet::format::_Format;
use crate::api::worksheet::WorkSheet;
use crate::Format;
use crate::result::WorkSheetResult;

pub trait Col: _Col {
    fn set_columns<R: LocationRange>(&mut self, col_range: R, column: &Column) -> WorkSheetResult<()> {
        self.set_by_column(col_range, &column)?;
        Ok(())
    }
    fn get_columns<R: LocationRange>(&self, col_range: R) -> WorkSheetResult<HashMap<String, Column>> {
        self.list_by_range(col_range)
    }
    fn set_columns_width<R: LocationRange>(&mut self, col_range: R, width: f64) -> WorkSheetResult<()> {
        let mut col_set = Column::default();
        col_set.width = Some(width);
        self.set_by_column(col_range, &col_set)?;
        Ok(())
    }
    fn get_columns_width<R: LocationRange>(&self, col_range: R) -> WorkSheetResult<HashMap<String, Option<f64>>> {
        let columns = self.list_by_range(col_range)?;
        let widths = columns.iter().map(|(k, v)| (k.clone(), v.width)).collect();
        Ok(widths)
    }
    fn set_columns_with_format<R: LocationRange>(&mut self, col_range: R, column: &Column, format: &Format) -> WorkSheetResult<()> {
        let mut col = column.clone();
        col.style = Some(self.add_format(format));
        self.set_by_column(col_range, &col)?;
        Ok(())
    }
    fn get_columns_with_format<R: LocationRange>(&self, col_range: R) -> WorkSheetResult<HashMap<String, (Column, Option<Format>)>> {
        let all = self.list_by_range(col_range)?
            .iter()
            .map(|(k, v)| (
                k.clone(),
                (*v,
                 match v.style {
                     None => None,
                     Some(style) => Some(self.get_format(style))
                 })
            ))
            .collect();
        Ok(all)
    }
    fn set_columns_width_pixels<R: LocationRange>(&mut self, col_range: R, width: f64) -> WorkSheetResult<()> {
        let mut col_set = Column::default();
        col_set.width = Some(0.5 * width);
        self.set_by_column(col_range, &col_set)?;
        Ok(())
    }
    fn set_columns_width_with_format<R: LocationRange>(&mut self, col_range: R, width: f64, format: &Format) -> WorkSheetResult<()> {
        let mut col_set = Column::default();
        col_set.style = Some(self.add_format(format));
        col_set.width = Some(width);
        self.set_by_column(col_range, &col_set)?;
        Ok(())
    }
    fn set_columns_width_pixels_with_format<R: LocationRange>(&mut self, col_range: R, width: f64, format: &Format) -> WorkSheetResult<()> {
        let mut col_set = Column::default();
        col_set.style = Some(self.add_format(format));
        col_set.width = Some(0.5 * width);
        self.set_by_column(col_range, &col_set)?;
        Ok(())
    }
    fn hide_columns<R: LocationRange>(&mut self, col_range: R) -> WorkSheetResult<()> {
        let mut col_set = Column::default();
        col_set.hidden = Some(1);
        self.set_by_column(col_range, &col_set)?;
        Ok(())
    }
    fn set_columns_level<R: LocationRange>(&mut self, col_range: R, level: u32) -> WorkSheetResult<()> {
        let mut col_set = Column::default();
        col_set.outline_level = Some(level);
        self.set_by_column(col_range, &col_set)?;
        Ok(())
    }
    fn collapse_columns<R: LocationRange>(&mut self, col_range: R) -> WorkSheetResult<()> {
        let mut col_set = Column::default();
        col_set.collapsed = Some(1);
        self.set_by_column(col_range, &col_set)?;
        Ok(())
    }
}

pub(crate) trait _Col: _Format {
    fn set_by_column<R: LocationRange>(&mut self, col_range: R, col_set: &Column) -> WorkSheetResult<()>;
    fn list_by_range<R: LocationRange>(&self, col_range: R) -> WorkSheetResult<HashMap<String, Column>>;
}

impl _Col for WorkSheet {
    fn set_by_column<R: LocationRange>(&mut self, col_range: R, col_set: &Column) -> WorkSheetResult<()> {
        self.worksheet.set_col_by_column(col_range, col_set)?;
        Ok(())
    }

    fn list_by_range<R: LocationRange>(&self, col_range: R) -> WorkSheetResult<HashMap<String, Column>> {
        let columns = self.worksheet.get_col(col_range)?;
        Ok(columns)
    }
}
