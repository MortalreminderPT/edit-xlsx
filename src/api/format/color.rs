//!
//! This module contains the [`FormatColor`] Enum, which used to set the color for the format.
//! [`FormatColor`] is mainly used for setting fonts, backgrounds, and sheet tab colors.
//! # Examples
//!
//! Use [`FormatColor`] to set font color
//!
//! ```
//! use edit_xlsx::{Format, FormatColor, Workbook, Write};
//! let mut workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
//! let worksheet = workbook.get_worksheet_mut_by_name("worksheet").unwrap();
//! let green = Format::default().set_color(FormatColor::RGB(0, 255, 0));
//! worksheet.write_with_format("B6", "Green text", &green).unwrap();
//! workbook.save_as("./examples/color_set_font_color.xlsx").unwrap();
//! ```
//!
//! Use [`FormatColor`] to set background color
//!
//! **NOTICE**: According to Excel file format standard,
//! when we are setting the background of xlsx,
//! we are just modifying the background color of all the columns in xlsx.
//!
//! And when you set a color for a column,
//! you must also set its width, or it will be hidden (i.e. width 0)
//!
//! ```
//! use edit_xlsx::{Column, Format, FormatColor, Workbook, WorkSheetCol, Write};
//! let mut workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
//! let worksheet = workbook.get_worksheet_mut_by_name("worksheet").unwrap();
//! let grey = Format::default().set_background_color(FormatColor::RGB(100, 100, 100));
//! worksheet.set_columns_width_with_format("A:XFD", 8.12, &grey).unwrap();
//! workbook.save_as("./examples/color_set_background_color.xlsx").unwrap();
//! ```
//!
//! Use [`FormatColor`] to set sheet tab color
//!
//! ```
//! use edit_xlsx::{Column, Format, FormatColor, Workbook, WorkSheetCol, Write};
//! let mut workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
//! let worksheet = workbook.get_worksheet_mut_by_name("worksheet").unwrap();
//! let yellow = FormatColor::RGB(255, 255, 0);
//! worksheet.set_tab_color(&yellow);
//! workbook.save_as("./examples/color_set_tab_color.xlsx").unwrap();
//! ```

use crate::xml::common::FromFormat;
use crate::xml::style::color::Color;

/// [`FormatColor`] is mainly used for setting fonts, backgrounds, and sheet tab colors.
///
/// # Fields:
/// | unit | fields |meaning|
/// | ---- | ---- |----|
/// | `Default` |  |Default, no color|
/// | `Index` | `index_id: u8,` | Using the [colorIndex property](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12)),<br/> `index_id`: color index id. |
/// | `Theme` | `theme_id: u8,`<br/>`tint: f64` |Use the Theme color format, i.e. leave the color decision to the current theme of xlsx.<br/>You can determine the theme color id and tint by `theme_id` and `tint` |
/// | `RGB` | `red: u8,`<br/>`green: u8,`<br/>`blue: u8,` |Using the RGB color format.|
#[derive(Clone, Debug, PartialEq, Copy)]
pub enum FormatColor {
    Default,
    Index(u8),
    Theme(u8, f64),
    RGB(u8, u8, u8),
}


impl Default for FormatColor {
    fn default() -> Self {
        Self::Default
    }
}

impl FromFormat<FormatColor> for Color {
    fn set_attrs_by_format(&mut self, format: &FormatColor) {
        match format {
            FormatColor::Default => *self = Color::default(),
            FormatColor::RGB(r, g, b) => *self = Color::from_rgb(*r, *g, *b),
            FormatColor::Index(id) => *self = Color::from_index(*id),
            FormatColor::Theme(theme, tint) => *self = Color::from_theme(*theme, *tint),
        }
    }

    fn set_format(&self, format: &mut FormatColor) {
        *format = if let Some(id) = self.indexed {
            FormatColor::Index(id)
        } else if let (Some(theme), tint) = (self.theme, self.tint) {
            FormatColor::Theme(theme, tint.unwrap_or_default())
        } else if let Some(color) = &self.rgb {
            let argb: Vec<u8> = color
                .chars()
                .collect::<Vec<char>>()
                .chunks(2)
                .map(|chunk| {
                    let hex_string: String = chunk.iter().collect();
                    u8::from_str_radix(&hex_string, 16).unwrap()
                })
                .collect();
            FormatColor::RGB(argb[1], argb[2], argb[3])
        } else {
            FormatColor::Default
        };
    }
}