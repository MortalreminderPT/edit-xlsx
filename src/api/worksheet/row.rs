use crate::Format;
use crate::api::worksheet::format::_Format;
use crate::api::worksheet::WorkSheet;
use crate::result::RowError::RowNotFound;
use crate::result::WorkSheetError::RowError;
use crate::result::WorkSheetResult;

///
/// Manages the Row settings that are modified,
/// only the parts that are not None are applied to the Row's modifications.
///
#[derive(Copy, Debug, Clone, Default)]
pub struct Row {
    pub height: Option<f64>,
    pub(crate) style: Option<u32>,
    pub outline_level: Option<u32>,
    pub hidden: Option<u8>,
    pub collapsed: Option<u8>,
}

pub trait WorkSheetRow: _Row {
    fn get_row(&self, row: u32) -> WorkSheetResult<Row> {
        self.get_by_row_number(row)
    }
    
    fn get_row_with_format(&self, row: u32) -> WorkSheetResult<(Row, Option<Format>)> {
        let row = self.get_by_row_number(row);
        match row {
            Ok(row) => {
                match row.style {
                    None => Ok((row, None)),
                    Some(style_id) => Ok((row, Some(self.get_format(style_id))))
                }
            }
            Err(_) => Err(RowError(RowNotFound))
        }
    }

    fn get_row_height(&self, row: u32) -> WorkSheetResult<Option<f64>> {
        let row = self.get_by_row_number(row)?;
        Ok(row.height)
    }

    ///
    /// set methods
    ///

    ///
    /// set row fields and format
    ///
    fn set_row(&mut self, row_number: u32, row: &Row) -> WorkSheetResult<()> {
        self.set_by_row(row_number, row)?;
        Ok(())
    }
    fn set_row_with_format(&mut self, row_number: u32, row: &Row, format: &Format) -> WorkSheetResult<()> {
        let mut row = row.clone();
        row.style = Some(self.add_format(format));
        self.set_by_row(row_number, &row)?;
        Ok(())
    }

    ///
    /// set row height and format
    ///
    fn set_row_height(&mut self, row: u32, height: f64) -> WorkSheetResult<()> {
        let mut row_set = Row::default();
        row_set.height = Some(height);
        self.set_by_row(row, &row_set)?;
        Ok(())
    }
    fn set_row_height_pixels(&mut self, row: u32, height: f64) -> WorkSheetResult<()> {
        let mut row_set = Row::default();
        row_set.height = Some(0.5 * height);
        self.set_by_row(row, &row_set)?;
        Ok(())
    }
    fn set_row_height_with_format(&mut self, row: u32, height: f64, format: &Format) -> WorkSheetResult<()> {
        let mut row_set = Row::default();
        row_set.height = Some(height);
        row_set.style = Some(self.add_format(format));
        self.set_by_row(row, &row_set)?;
        Ok(())
    }
    fn set_row_height_pixels_with_format(&mut self, row: u32, height: f64, format: &Format) -> WorkSheetResult<()> {
        let mut row_set = Row::default();
        row_set.height = Some(0.5 * height);
        row_set.style = Some(self.add_format(format));
        self.set_by_row(row, &row_set)?;
        Ok(())
    }

    ///
    /// hide row
    ///
    fn hide_row(&mut self, row: u32) -> WorkSheetResult<()> {
        let mut row_set = Row::default();
        row_set.hidden = Some(1);
        self.set_by_row(row, &row_set)?;
        Ok(())
    }

    ///
    /// collapse rows
    ///
    fn set_row_level(&mut self, row: u32, level: u32) -> WorkSheetResult<()> {
        let mut row_set = Row::default();
        row_set.outline_level = Some(level);
        self.set_by_row(row, &row_set)?;
        Ok(())
    }

    fn collapse_row(&mut self, row: u32) -> WorkSheetResult<()> {
        let mut row_set = Row::default();
        row_set.collapsed = Some(1);
        self.set_by_row(row, &row_set)?;
        Ok(())
    }
}

pub(crate) trait _Row: _Format {
    fn set_by_row(&mut self, row: u32, row_set: &Row) -> WorkSheetResult<()>;
    fn get_by_row_number(&self, row_number: u32) -> WorkSheetResult<Row>;
    // fn get_custom_row_height(&self, row: u32) -> WorkSheetResult<f64>;
}

impl _Row for WorkSheet {
    fn set_by_row(&mut self, row: u32, row_set: &Row) -> WorkSheetResult<()> {
        self.worksheet.sheet_data.set_by_row(row, row_set);
        Ok(())
    }

    fn get_by_row_number(&self, row_number: u32) -> WorkSheetResult<Row> {
        let row = self.worksheet.sheet_data.get_api_row(row_number)?;
        Ok(row)
    }

    // fn get_custom_row_height(&self, row: u32) -> WorkSheetResult<f64> {
    //     let custiom_height = self.worksheet.sheet_data.get_row_height(row)?;
    //     Ok(custiom_height)
    // }
}