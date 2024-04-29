//!
//! This module contains the [`FormatAlign`] struct, which used to edit the align style of the [`Cell`],
//! including set vertical and horizontal align, reading order and indent.
//!
//! # Examples
//!
//! Set vertical and horizontal align
//! ```
//! use edit_xlsx::{Cell, Format, FormatAlign, FormatAlignType, Workbook, Write};
//! // Set cell text alignment to top left
//! let format = Format::default()
//!     .set_align(FormatAlignType::Top)
//!     .set_align(FormatAlignType::Left);
//! // Or use
//! // let mut format = Format::default();
//! // let mut format_align = FormatAlign::default();
//! // format_align.vertical = Some(FormatAlignType::Top);
//! // format_align.horizontal = Some(FormatAlignType::Left);
//! // format.align = format_align;
//! let mut workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
//! let worksheet = workbook.get_worksheet_mut_by_name("worksheet").unwrap();
//! for row in 6..=15 {
//!     for col in 3..=12 {
//!         worksheet.write_with_format((row, col), 20.0, &format).unwrap()
//!     }
//! }
//! workbook.save_as("./examples/align.xlsx").unwrap();
//! ```
//!
//! Set reading order
//! ```
//! use edit_xlsx::{Cell, Format, FormatAlign, FormatAlignType, Workbook, Write};
//! let format = Format::default()
//!     .set_reading_order(2);
//! let mut workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
//! let worksheet = workbook.get_worksheet_mut_by_name("worksheet").unwrap();
//! for row in 6..=15 {
//!     for col in 3..=12 {
//!         worksheet.write_with_format((row, col), "نص عربي / English text", &format).unwrap();
//!     }
//! }
//! workbook.save_as("./examples/reading_order.xlsx").unwrap();
//! ```
//!
//! Set indent
//! ```
//! use edit_xlsx::{Cell, Format, FormatAlign, FormatAlignType, Workbook, Write};
//! let mut format = Format::default()
//!     .set_align(FormatAlignType::Top)
//!     .set_align(FormatAlignType::Left);
//! let mut workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
//! let worksheet = workbook.get_worksheet_mut_by_name("worksheet").unwrap();
//! for col in 3..=12 {
//!     let mut indent = 0;
//!     for row in 6..=15 {
//!         let format = format.clone().set_indent(indent);
//!         worksheet.write_with_format((row, col), 10.0, &format).unwrap();
//!         indent += 1;
//!     }
//! }
//! workbook.save_as("./examples/indent.xlsx").unwrap();
//! ```
use crate::Cell;
use crate::xml::common::FromFormat;

///
/// [`FormatAlign`] struct, which used to edit the align style of the [`Cell`],
/// including set vertical and horizontal align(by [`FormatAlignType`]), reading order and indent.
/// # Fields
/// | field        | type        | meaning                                                      |
/// | ------------ | ----------- | ------------------------------------------------------------ |
/// | `horizontal`     | [`Option<FormatAlignType>`] | The [`Cell`]'s horizontal alignment.<br> It's Some value can be [`FormatAlignType`] Left, Right, Center |
/// | `vertical`    | [`Option<FormatAlignType>`] | The [`Cell`]'s vertical alignment.<br> It's Some value can be [`FormatAlignType`] Top, Bottom, VerticalCenter |
/// | `reading_order`      | u8 | The [`Cell`]'s reading order<br>1: left to right<br>2: right to left|
/// | `indent`   | u8 | The [`Cell`]'s indent               |
#[derive(Clone, Debug, PartialEq)]
pub struct FormatAlign {
    pub horizontal: Option<FormatAlignType>,
    pub vertical: Option<FormatAlignType>,
    pub reading_order: u8,
    pub indent: u8,
}

impl Default for FormatAlign {
    fn default() -> Self {
        Self {
            horizontal: None,
            vertical: None,
            reading_order: 1,
            indent: 0,
        }
    }
}

/// [`FormatAlignType`] determines the position of the alignment.
///
/// # Fields:
/// | unit | meaning|
/// | ---- | ---- |
/// | `Top` |top-align|
/// | `Center` |horizontal center-align|
/// | `Bottom` |bottom-align|
/// | `Left` |left-align|
/// | `VerticalCenter` |vertical center-align|
/// | `Right` |right-align|
#[derive(Copy, Clone, Debug, PartialEq)]
pub enum FormatAlignType {
    Top,
    Center,
    Bottom,
    Left,
    VerticalCenter,
    Right,
}

impl Default for FormatAlignType {
    fn default() -> Self {
        FormatAlignType::Center
    }
}

impl FormatAlignType {
    pub(crate) fn to_str(&self) -> &str {
        match self {
            FormatAlignType::Top => "top",
            FormatAlignType::Center => "center",
            FormatAlignType::Bottom => "bottom",
            FormatAlignType::Left => "left",
            FormatAlignType::VerticalCenter => "center",
            FormatAlignType::Right => "right",
        }
    }

    pub(crate) fn from_str(format_align_type: Option<&String>, is_horizontal: bool) -> Option<FormatAlignType> {
        if let Some(format_align_type) = format_align_type {
            let format_align_type = format_align_type.as_str();
            Some(match format_align_type {
                "top" => FormatAlignType::Top,
                "center" => if is_horizontal { FormatAlignType::Center } else { FormatAlignType::VerticalCenter },
                "bottom" => FormatAlignType::Bottom,
                "left" => FormatAlignType::Left,
                "right" => FormatAlignType::Right,
                _ => FormatAlignType::default()
            })
        } else {
            None
        }
    }
}

impl FromFormat<FormatAlignType> for String {
    fn set_attrs_by_format(&mut self, format: &FormatAlignType) {
        *self = format.to_str().to_string();
    }

    fn set_format(&self, format: &mut FormatAlignType) {
        todo!()
    }
}
