//!
//! Manages the Column settings that are modified,
//! only the parts that are not None are applied to the [`Column`]'s modifications.
//!
//! This module contains the [`Column`] type, the [`WorkSheetCol`] trait for
//! [`WorkSheet`]s working with [`Column`]s.
//!
//! # Examples
//!
//! There are multiple ways to create a new [`Column`]
//!
//! ```
//! use edit_xlsx::Column;
//! let col = Column::new(15.0, 2, 1, 0);
//! assert_eq!(col.width, Some(15.0));
//! assert_eq!(col.outline_level, Some(2));
//! assert_eq!(col.hidden, Some(1));
//! assert_eq!(col.collapsed, Some(0));
//! ```
//! Since the [`Column`] records the fields you want to update to worksheets, all their fields are optional
//! If you create a new [`Column`] by default, the fields in it will be filled with None.
//! ```
//! use edit_xlsx::Column;
//! let col = Column::default();
//! assert_eq!(col.width, None);
//! assert_eq!(col.outline_level, None);
//! assert_eq!(col.hidden, None);
//! assert_eq!(col.collapsed, None);
//! ```
//!
//! You can update Worksheet Columns by using the methods in WorkSheetCol
//! ```
//! use edit_xlsx::{Column, Workbook, WorkSheetCol};
//! let mut workbook = Workbook::new();
//! let worksheet = workbook.get_worksheet_mut(1).unwrap();
//! let col = Column::new(15.0, 2, 1, 0);
//! worksheet.set_columns("A:C", &col).unwrap();
//! workbook.save_as("./examples/col_update_columns.xlsx").unwrap()
//! ```

use std::collections::HashMap;
pub(crate) use crate::api::cell::location::LocationRange;
use crate::api::worksheet::format::_Format;
use crate::api::worksheet::WorkSheet;
use crate::{Format, Row, Cell};
use crate::result::WorkSheetResult;

/// [`Column`] records the fields you want to update to worksheets.
///
/// # Fields:
/// | field | type |meaning|
/// | ---- | ---- |----|
/// | `width` | `Option<f64>` |The custom width you want to update with.|
/// | `outline_level` | `Option<u8>` |The outline level of a column, learn more from [official documentation](https://support.microsoft.com/en-us/office/outline-group-data-in-a-worksheet-08ce98c4-0063-4d42-8ac7-8278c49e9aff).|
/// | `hidden` | `Option<u8>` |Whether the column is hidden or not.|
/// | `collapsed` | `Option<u8>` |collapse columns to group them.|
#[derive(Copy, Clone, Default, Debug)]
pub struct Column {
    pub width: Option<f64>,
    pub(crate) style: Option<u32>,
    pub outline_level: Option<u8>,
    pub hidden: Option<u8>,
    pub collapsed: Option<u8>,
}

impl Column {
    //
    // constructors
    //
    /// If you need to customize each field, you can use the [`Column::new()`] method to create a [`Column`]
    /// ```
    /// use edit_xlsx::Column;
    /// let col = Column::new(15.0, 2, 1, 0);
    /// assert_eq!(col.width, Some(15.0));
    /// assert_eq!(col.outline_level, Some(2));
    /// assert_eq!(col.hidden, Some(1));
    /// assert_eq!(col.collapsed, Some(0));
    /// ```
    pub fn new(width: f64, outline_level: u8, hidden: u8, collapsed: u8) -> Column {
        Column {
            width: Some(width),
            style: None,
            outline_level: Some(outline_level),
            hidden: Some(hidden),
            collapsed: Some(collapsed),
        }
    }

    /// If you want to custom the format of col, you can use [`Column::new_by_worksheet()`] method.
    /// NOTICE: A [`Column`] created using the [`Column::new_by_worksheet()`] method can only be used in incoming worksheets.
    /// ```
    /// use edit_xlsx::{Workbook, WorkSheetCol, Column, Format, FormatColor};
    /// let red = Format::default().set_background_color(FormatColor::RGB(255, 0, 0));
    /// let mut workbook = Workbook::new();
    /// let worksheet = workbook.get_worksheet_mut(1).unwrap();
    /// let col = Column::new_by_worksheet(15.0, 2, 1, 0, &red, worksheet);
    /// worksheet.set_columns("A:C", &col).unwrap();
    /// workbook.save_as("./examples/col_new_by_worksheet.xlsx").unwrap()
    /// ```
    pub fn new_by_worksheet(width: f64, outline_level: u8
                            , hidden: u8, collapsed: u8
                            , format: &Format, work_sheet: &mut WorkSheet) -> Column {
        let mut col = Column::new(width, outline_level, hidden, collapsed);
        col.style = Some(work_sheet.add_format(format));
        col
    }
}

/// [`WorkSheetCol`] is a trait for [`WorkSheet`]s that allowing them working with [`Column`]s.
///
/// Not only does it support reading and updating cols directly,
/// but it also provides a set of suggested methods for reading and updating cols swiftly.
pub trait WorkSheetCol: _Col {
    //
    // get methods
    //

    /// Get the [`Column`]s based on the col range,
    /// note that the col range starts with 1.
    /// # Example
    /// ```
    /// use edit_xlsx::{Workbook, WorkSheetCol};
    /// let workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
    /// let worksheet = workbook.get_worksheet_by_name("worksheet").unwrap();
    /// let columns = worksheet.get_columns("B:B").unwrap();
    /// let first_col = columns.get("B:B").unwrap();
    /// // Convert to u32 to reduce error
    /// assert_eq!(first_col.width.unwrap() as u32, 26);
    /// ```
    fn get_columns<R: LocationRange>(&self, col_range: R) -> WorkSheetResult<HashMap<String, Column>> {
        self.list_by_range(col_range)
    }

    /// Get the [`Column`]s and [`Format`]s of cols based on the col range,
    /// note that the col range starts with 1.
    /// # Example
    /// ```
    /// use edit_xlsx::{FormatColor, Workbook, WorkSheetCol};
    /// let workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
    /// let worksheet = workbook.get_worksheet_by_name("worksheet").unwrap();
    /// let columns_with_format = worksheet.get_columns_with_format("M:N").unwrap();
    /// let (_, format) = columns_with_format.get("M:N").unwrap();
    /// let format_clone = format.clone();
    /// assert_eq!(format_clone.unwrap().get_background().fg_color, FormatColor::RGB(255, 0, 0));
    /// ```
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

    /// Get the custom width of cols based on the col range,
    /// note that the col range starts with 1.
    /// # Example
    /// ```
    /// use edit_xlsx::{Workbook, WorkSheetCol};
    /// let workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
    /// let worksheet = workbook.get_worksheet_by_name("worksheet").unwrap();
    /// let widths = worksheet.get_columns_width("A:C").unwrap();
    /// widths.iter().for_each(|(k, v)|{
    ///   //Convert to u32 to reduce error
    ///   let width = v.unwrap() as u32;
    ///   match k.as_str() {
    ///     "A:A"|"C:C" => assert_eq!(width, 12),
    ///     "B:B" => assert_eq!(width, 26),
    ///     _ => {}
    ///   }
    /// });
    /// ```
    fn get_columns_width<R: LocationRange>(&self, col_range: R) -> WorkSheetResult<HashMap<String, Option<f64>>> {
        let columns = self.list_by_range(col_range)?;
        let widths = columns.iter().map(|(k, v)| (k.clone(), v.width)).collect();
        Ok(widths)
    }

    //
    // set methods
    //

    /// update columns by [`Column`].
    ///
    /// Only not none fields will be updated
    ///
    /// # Example
    /// ```
    /// use edit_xlsx::{Column, Workbook, WorkSheetCol};
    /// let mut workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
    /// let mut worksheet = workbook.get_worksheet_mut_by_name("worksheet").unwrap();
    /// let mut col = Column::default();
    /// col.outline_level = Some(1);
    /// col.hidden = Some(1);
    /// worksheet.set_columns("B:E", &col).unwrap();
    /// workbook.save_as("./examples/col_set_columns.xlsx").unwrap()
    /// ```
    fn set_columns<R: LocationRange>(&mut self, col_range: R, column: &Column) -> WorkSheetResult<()> {
        self.set_by_column(col_range, column)?;
        Ok(())
    }

    /// update columns and formats by [`Column`].
    ///
    /// Only not none fields will be updated.
    ///
    /// NOTICE: Changing the [`Column`]'s [`Format`] does not mean that the effect can be seen directly in Excel,
    /// because the style priority is [`Cell`]>[`Row`]>[`Column`].
    /// # Example
    /// ```
    /// use edit_xlsx::{Column, Format, FormatColor, Workbook, WorkSheetCol};
    /// let mut workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
    /// let mut worksheet = workbook.get_worksheet_mut_by_name("worksheet").unwrap();
    /// let mut col = Column::default();
    /// col.outline_level = Some(1);
    /// col.hidden = Some(1);
    /// let white_font = Format::default().set_background_color(FormatColor::Index(4)).set_color(FormatColor::RGB(255, 255, 255));
    /// worksheet.set_columns_with_format("B:E", &col, &white_font).unwrap();
    /// workbook.save_as("./examples/col_set_columns_with_format.xlsx").unwrap()
    /// ```
    fn set_columns_with_format<R: LocationRange>(&mut self, col_range: R, column: &Column, format: &Format) -> WorkSheetResult<()> {
        let mut col = column.clone();
        col.style = Some(self.add_format(format));
        self.set_by_column(col_range, &col)?;
        Ok(())
    }

    /// set the width of columns by columns range,
    /// The effect is the same as
    /// # Basic Example
    /// ```
    /// use edit_xlsx::Column;
    /// let mut col = Column::default();
    /// col.width = Some(8.0);
    /// // worksheet.set_columns("A:C", &col);
    /// ```
    fn set_columns_width<R: LocationRange>(&mut self, col_range: R, width: f64) -> WorkSheetResult<()> {
        let mut col_set = Column::default();
        col_set.width = Some(width);
        self.set_by_column(col_range, &col_set)?;
        Ok(())
    }
    fn set_columns_width_pixels<R: LocationRange>(&mut self, col_range: R, width: f64) -> WorkSheetResult<()> {
        let mut col_set = Column::default();
        col_set.width = Some(0.5 * width);
        self.set_by_column(col_range, &col_set)?;
        Ok(())
    }

    /// set the width of columns by columns range,
    /// The effect is the same as
    /// # Basic Example
    /// ```
    /// use edit_xlsx::{Column, Format};
    /// let mut col = Column::default();
    /// col.width = Some(8.0);
    /// let format = Format::default();
    /// // worksheet.set_columns_with_format("A:C", &col, &format);
    /// ```
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

    /// hide columns by column range,
    /// The effect is the same as
    /// # Basic Example
    /// ```
    /// use edit_xlsx::Column;
    /// let mut col = Column::default();
    /// col.hidden = Some(1);
    /// // worksheet.set_columns("A:C", &col);
    /// ```
    fn hide_columns<R: LocationRange>(&mut self, col_range: R) -> WorkSheetResult<()> {
        let mut col_set = Column::default();
        col_set.hidden = Some(1);
        self.set_by_column(col_range, &col_set)?;
        Ok(())
    }

    /// set the outline of columns by column range,
    /// The effect is the same as
    /// # Basic Example
    /// ```
    /// use edit_xlsx::Column;
    /// let mut col = Column::default();
    /// col.outline_level = Some(1);
    /// // worksheet.set_columns("A:C", &col);
    /// ```
    fn set_columns_level<R: LocationRange>(&mut self, col_range: R, level: u8) -> WorkSheetResult<()> {
        let mut col_set = Column::default();
        col_set.outline_level = Some(level);
        self.set_by_column(col_range, &col_set)?;
        Ok(())
    }
    /// collapse columns by column range,
    /// The effect is the same as
    /// # Basic Example
    /// ```
    /// use edit_xlsx::Column;
    /// let mut col = Column::default();
    /// col.collapsed = Some(1);
    /// // worksheet.set_columns("A:C", &col);
    /// ```
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
