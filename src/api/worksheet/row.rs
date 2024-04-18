//!
//! Manages the Row settings that are modified,
//! only the parts that are not None are applied to the [`Row`]'s modifications.
//!
//! This module contains the [`Row`] type, the [`WorkSheetRow`] trait for
//! [`WorkSheet`]s working with [`Row`]s.
//!
//! # Examples
//!
//! There are multiple ways to create a new [`Row`]
//!
//! ```
//! use edit_xlsx::Row;
//! let row = Row::new(15.0, 2, 1, 0);
//! assert_eq!(row.height, Some(15.0));
//! assert_eq!(row.outline_level, Some(2));
//! assert_eq!(row.hidden, Some(1));
//! assert_eq!(row.collapsed, Some(0));
//! ```
//! Since the [`Row`] records the fields you want to update to worksheets, all their fields are optional
//! If you create a new [`Row`] by default, the fields in it will be filled with None.
//! ```
//! use edit_xlsx::Row;
//! let row = Row::default();
//! assert_eq!(row.height, None);
//! assert_eq!(row.outline_level, None);
//! assert_eq!(row.hidden, None);
//! assert_eq!(row.collapsed, None);
//! ```
//!
//! You can update Worksheet rows by using the methods in WorkSheetRow
//! ```
//! use edit_xlsx::{Row, Workbook, WorkSheetRow};
//! let mut workbook = Workbook::new();
//! let worksheet = workbook.get_worksheet_mut(1).unwrap();
//! let row = Row::new(15.0, 2, 1, 0);
//! worksheet.set_row(1, &row).unwrap();
//! workbook.save_as("./examples/row_update_row.xlsx").unwrap()
//! ```

use crate::Format;
use crate::api::worksheet::format::_Format;
use crate::api::worksheet::WorkSheet;
use crate::result::RowError::RowNotFound;
use crate::result::WorkSheetError::RowError;
use crate::result::WorkSheetResult;

/// [`Row`] records the fields you want to update to worksheets.
///
/// # Fields:
/// | field | type |meaning|
/// | ---- | ---- |----|
/// | `height` | `Option<f64>` |The customed height you want to update with.|
/// | `outline_level` | `Option<u8>` |The ouline level of a row, learn more from [official documentation](https://support.microsoft.com/en-us/office/outline-group-data-in-a-worksheet-08ce98c4-0063-4d42-8ac7-8278c49e9aff).|
/// | `hidden` | `Option<u8>` |Whether the row is hidden or not.|
/// | `collapsed` | `Option<u8>` |collapse rows to group them.|
#[derive(Copy, Debug, Clone, Default)]
pub struct Row {
    pub height: Option<f64>,
    pub(crate) style: Option<u32>,
    pub outline_level: Option<u8>,
    pub hidden: Option<u8>,
    pub collapsed: Option<u8>,
}

impl Row {
    //
    // constructors
    //
    /// If you need to customize each field, you can use the [`Row::new()`] method to create a [`Row`]
    /// ```
    /// use edit_xlsx::Row;
    /// let row = Row::new(15.0, 2, 1, 0);
    /// assert_eq!(row.height, Some(15.0));
    /// assert_eq!(row.outline_level, Some(2));
    /// assert_eq!(row.hidden, Some(1));
    /// assert_eq!(row.collapsed, Some(0));
    /// ```
    pub fn new(height: f64, outline_level: u8, hidden: u8, collapsed: u8) -> Row {
        Row {
            height: Some(height),
            style: None,
            outline_level: Some(outline_level),
            hidden: Some(hidden),
            collapsed: Some(collapsed),
        }
    }

    /// If you want to custom the format of row, you can use [`Row::new_by_worksheet()`] method.
    /// NOTICE: A [`Row`] created using the [`Row::new_by_worksheet()`] method can only be used in incoming worksheets.
    /// ```
    /// use edit_xlsx::{Workbook, WorkSheetRow, Row, Format, FormatColor};
    /// let red = Format::default().set_background_color(FormatColor::RGB(255, 0, 0));
    /// let mut workbook = Workbook::new();
    /// let worksheet = workbook.get_worksheet_mut(1).unwrap();
    /// let row = Row::new_by_worksheet(15.0, 2, 1, 0, &red, worksheet);
    /// worksheet.set_row(1, &row).unwrap();
    /// workbook.save_as("./examples/row_new_by_worksheet.xlsx").unwrap()
    /// ```
    pub fn new_by_worksheet(height: f64, outline_level: u8
                        , hidden: u8, collapsed: u8
                        , format: &Format, work_sheet: &mut WorkSheet) -> Row {
        let mut row = Row::new(height, outline_level, hidden, collapsed);
        row.style = Some(work_sheet.add_format(format));
        row
    }
}

/// [`WorkSheetRow`] is a trait for [`WorkSheet`]s that allowing them working with [`Row`]s.
///
/// Not only does it support reading and updating rows directly,
/// but it also provides a set of suggested methods for reading and updating rows swiftly.
pub trait WorkSheetRow: _Row {
    //
    // get methods
    //

    /// Get the [`Row`] of a row based on the row number,
    /// note that the row number starts with 1.
    /// # Example
    /// ```
    /// use edit_xlsx::{Workbook, WorkSheetRow};
    /// let workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
    /// let worksheet = workbook.get_worksheet_by_name("worksheet").unwrap();
    /// let first_row = worksheet.get_row(1).unwrap();
    /// assert_eq!(first_row.height.unwrap(), 28.25);
    /// ```
    fn get_row(&self, row: u32) -> WorkSheetResult<Row> {
        self.get_by_row_number(row)
    }

    /// Get the [`Row`] and [`Format`] of a row based on the row number,
    /// note that the row number starts with 1.
    /// # Example
    /// ```
    /// use edit_xlsx::{Workbook, WorkSheetRow};
    /// let workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
    /// let worksheet = workbook.get_worksheet_by_name("worksheet").unwrap();
    /// let (_, third_row_format) = worksheet.get_row_with_format(3).unwrap();
    /// assert_eq!(third_row_format.unwrap().font.bold, true);
    /// ```
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

    /// Get the custom height of a row based on the row number,
    /// note that the row number starts with 1.
    /// Only rows have custom height can return their height.
    /// # Example
    /// ```
    /// use edit_xlsx::{Workbook, WorkSheetRow};
    /// let workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
    /// let worksheet = workbook.get_worksheet_by_name("worksheet").unwrap();
    /// let first_height = worksheet.get_row_height(1).unwrap();
    /// let forth_height = worksheet.get_row_height(4).unwrap();
    /// assert_eq!(first_height, Some(28.25));
    /// assert_eq!(forth_height, None);
    /// ```
    fn get_row_height(&self, row: u32) -> WorkSheetResult<Option<f64>> {
        let row = self.get_by_row_number(row)?;
        Ok(row.height)
    }

    //
    // set methods
    //

    ///
    /// update a row by [`Row`], note that the row number starts with 1.
    ///
    /// Only not none fields will be updated
    ///
    /// # Example
    /// ```
    /// use edit_xlsx::{Row, Workbook, WorkSheetRow};
    /// let mut workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
    /// let mut worksheet = workbook.get_worksheet_mut_by_name("worksheet").unwrap();
    /// let mut row = Row::default();
    /// row.outline_level = Some(1);
    /// row.hidden = Some(1);
    /// worksheet.set_row(1, &row).unwrap();
    /// worksheet.set_row(2, &row).unwrap();
    /// worksheet.set_row(3, &row).unwrap();
    /// workbook.save_as("./examples/row_set_row.xlsx").unwrap()
    fn set_row(&mut self, row_number: u32, row: &Row) -> WorkSheetResult<()> {
        self.set_by_row(row_number, row)?;
        Ok(())
    }

    ///
    /// update a row by [`Row`], note that the row number starts with 1.
    ///
    /// Only not none fields will be updated
    ///
    /// # Example
    /// ```
    /// use edit_xlsx::{Format, FormatColor, Row, Workbook, WorkSheetRow};
    /// let mut workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
    /// let mut worksheet = workbook.get_worksheet_mut_by_name("worksheet").unwrap();
    /// let mut row = Row::default();
    /// row.height = Some(10.0);
    /// let red_font = Format::default().set_background_color(FormatColor::Index(4)).set_color(FormatColor::RGB(255, 0, 0));
    /// worksheet.set_row_with_format(6, &row, &red_font).unwrap();
    /// worksheet.set_row_with_format(7, &row, &red_font).unwrap();
    /// worksheet.set_row_with_format(8, &row, &red_font).unwrap();
    /// workbook.save_as("./examples/row_set_row_with_format.xlsx").unwrap()
    fn set_row_with_format(&mut self, row_number: u32, row: &Row, format: &Format) -> WorkSheetResult<()> {
        let mut row = row.clone();
        row.style = Some(self.add_format(format));
        self.set_by_row(row_number, &row)?;
        Ok(())
    }

    /// set the height of a row by row number,
    /// The effect is the same as
    /// # Basic Example
    /// ```
    /// use edit_xlsx::Row;
    /// let mut row = Row::default();
    /// row.height = Some(15.0);
    /// // worksheet.set_row(1, &row);
    /// ```
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

    /// set the height and format of a row by row number,
    /// The effect is the same as
    /// # Basic Example
    /// ```
    /// use edit_xlsx::{Format, Row};
    /// let mut row = Row::default();
    /// row.height = Some(15.0);
    /// let format = Format::default();
    /// // worksheet.set_row_with_format(1, &row, &format);
    /// ```
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

    /// hide a row by row number,
    /// The effect is the same as
    /// # Basic Example
    /// ```
    /// use edit_xlsx::Row;
    /// let mut row = Row::default();
    /// row.hidden = Some(1);
    /// // worksheet.set_row(1, &row);
    /// ```
    fn hide_row(&mut self, row: u32) -> WorkSheetResult<()> {
        let mut row_set = Row::default();
        row_set.hidden = Some(1);
        self.set_by_row(row, &row_set)?;
        Ok(())
    }

    /// set the outline of a row by row number,
    /// The effect is the same as
    /// # Basic Example
    /// ```
    /// use edit_xlsx::Row;
    /// let mut row = Row::default();
    /// row.outline_level = Some(1);
    /// // worksheet.set_row(1, &row);
    /// ```
    fn set_row_level(&mut self, row: u32, level: u8) -> WorkSheetResult<()> {
        let mut row_set = Row::default();
        row_set.outline_level = Some(level);
        self.set_by_row(row, &row_set)?;
        Ok(())
    }

    /// collapse a row by row number,
    /// The effect is the same as
    /// # Basic Example
    /// ```
    /// use edit_xlsx::Row;
    /// let mut row = Row::default();
    /// row.collapsed = Some(1);
    /// // worksheet.set_row(1, &row);
    /// ```
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
}