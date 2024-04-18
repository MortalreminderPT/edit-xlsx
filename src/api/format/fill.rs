//!
//! This module contains the [`FormatFill`] struct, which used to set the fill of a format.
//! [`FormatFill`] is used for fields that having both foreground color and background color,
//! mainly applied in the [`Row`], [`Column`] and [`Cell`]'s [`Format`] fill color.
//! # Examples
//!
//! Use [`FormatFill`] to fill columns' color
//! ```
//! use edit_xlsx::{Format, FormatColor, FormatFill, Workbook, WorkSheetCol, Write};
//! let mut workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
//! let worksheet = workbook.get_worksheet_mut_by_name("worksheet").unwrap();
//! let mut white = Format::default();
//! white.fill.fg_color = FormatColor::RGB(255, 255, 255);
//! white.fill.bg_color = FormatColor::RGB(0, 0, 200);
//! white.fill.pattern_type = "lightGrid".to_string();
//! worksheet.set_columns_width_with_format("A:XFD", 8.12, &white).unwrap();
//! workbook.save_as("./examples/fill_fill_columns_color.xlsx").unwrap();
//! ```
//!
//! Use [`FormatFill`] to fill row's color
//!
//! ```
//! use edit_xlsx::{Format, FormatColor, FormatFill, Row, Workbook, WorkSheetCol, WorkSheetRow, Write};
//! let mut workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
//! let worksheet = workbook.get_worksheet_mut_by_name("worksheet").unwrap();
//! let mut white = Format::default();
//! white.fill.fg_color = FormatColor::RGB(255, 255, 255);
//! white.fill.bg_color = FormatColor::RGB(200, 0, 0);
//! white.fill.pattern_type = "lightGrid".to_string();
//! worksheet.set_row_height_with_format(1, 15.0, &white).unwrap();
//! worksheet.set_row_height_with_format(2, 15.0, &white).unwrap();
//! worksheet.set_row_height_with_format(3, 15.0, &white).unwrap();
//! workbook.save_as("./examples/fill_fill_rows_color.xlsx").unwrap();
//! ```
//!
//! Use [`FormatFill`] to fill cell's color
//!
//! ```
//! use edit_xlsx::{Format, FormatColor, FormatFill, Workbook, Write};
//! let mut workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
//! let worksheet = workbook.get_worksheet_mut_by_name("worksheet").unwrap();
//! let mut yellow = Format::default();
//! yellow.fill.fg_color = FormatColor::RGB(255, 255, 0);
//! yellow.fill.bg_color = FormatColor::RGB(0, 255, 0);
//! yellow.fill.pattern_type = "lightGrid".to_string();
//! worksheet.write_with_format("B6", "Yellow fg and green bg", &yellow).unwrap();
//! workbook.save_as("./examples/fill_fill_cells_color.xlsx").unwrap();
//! ```


use crate::FormatColor;
use crate::xml::common::FromFormat;
use crate::xml::style::color::Color;
use crate::xml::style::fill::Fill;
use crate::{Row, Column, Cell, Format};

///
/// [`FormatFill`] is used for fields that having both foreground color and background color,
/// It's mainly applied in the [`Row`], [`Column`] and [`Cell`]'s [`Format`] fill color.
/// # Fields
/// | field        | type        | meaning                                                      |
/// | ------------ | ----------- | ------------------------------------------------------------ |
/// | `pattern_type` | [`String`]      | The color filling method, the contents of which can be referred to the [official documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.patternvalues?view=openxml-3.0.1) |
/// | `fg_color`     | [`FormatColor`] | The foreground color of filling                              |
/// | `bg_color`     | [`FormatColor`] | The background color of filling                              |
///
#[derive(Clone, Debug, PartialEq)]
pub struct FormatFill {
    pub pattern_type: String,
    pub fg_color: FormatColor,
    pub bg_color: FormatColor
}

impl Default for FormatFill {
    ///
    /// Method [`FormatFill::default()`] creates a blank [`FormatFill`] with no foreground color,
    /// a blank background color, and a `none` pattern_type.
    ///
    /// When using [`Format::default()`], its `fill` field will be filled by this method.
    fn default() -> Self {
        FormatFill {
            pattern_type: "none".to_string(),
            fg_color: FormatColor::default(),
            bg_color: FormatColor::Index(65),
        }
    }
}

impl FromFormat<FormatFill> for Fill {
    fn set_attrs_by_format(&mut self, format: &FormatFill) {
        self.pattern_fill.fg_color = Color::from_format(&format.fg_color);
        self.pattern_fill.bg_color = Color::from_format(&format.bg_color);
        self.pattern_fill.pattern_type = String::from(&format.pattern_type);
    }

    fn set_format(&self, format: &mut FormatFill) {
        format.fg_color = self.pattern_fill.fg_color.get_format();
        format.bg_color = self.pattern_fill.bg_color.get_format();
        format.pattern_type = self.pattern_fill.pattern_type.to_string();
    }
}