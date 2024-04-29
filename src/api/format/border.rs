//!
//! This module contains the [`FormatBorder`] struct,
//! which used to edit the border style of the [`Cell`] in each direction.
//!
//! # Examples
//!
//! Wrap cells
//! ```
//! use std::cell;
//! use edit_xlsx::{Cell, Format, FormatBorder, FormatBorderElement, FormatBorderType, FormatColor, Workbook, Write};
//! let format = Format::default().set_border(FormatBorderType::Thin);
//! // Or use
//! // let mut format = Format::default();
//! // let mut format_border = FormatBorder::default();
//! // let format_border_element = FormatBorderElement::from_border_type(&FormatBorderType::Thin);
//! // format_border.left = format_border_element;
//! // format_border.right = format_border_element;
//! // format_border.top = format_border_element;
//! // format_border.bottom = format_border_element;
//! // format.border = format_border;
//! let mut workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
//! let worksheet = workbook.get_worksheet_mut_by_name("worksheet").unwrap();
//! let mut cell: Cell<String> = Cell::default();
//! cell.format = Some(format);
//! for row in 6..=15 {
//!     for col in 3..=12 {
//!         worksheet.write_cell((row, col), &cell).unwrap()
//!     }
//! }
//! workbook.save_as("./examples/border_wrap_cells.xlsx").unwrap();
//! ```
//!
//! Use diagonal to creat a table
//! ```
//! use std::cell;
//! use edit_xlsx::{Cell, Format, FormatBorder, FormatBorderElement, FormatBorderType, FormatColor, Read, Workbook, Write};
//! let mut workbook = Workbook::new();
//! let worksheet = workbook.get_worksheet_mut(1).unwrap();
//! let mut cell: Cell<String> = Cell::default();
//! let mut format = Format::default();
//! format.border.diagonal = FormatBorderElement::from_border_type(&FormatBorderType::Thin);
//! cell.format = Some(format);
//! worksheet.write_cell("A1", &cell).unwrap();
//! // todo bug fix
//! worksheet.write_column("A2", &[1, 2, 3]).unwrap();
//! worksheet.write_column("B1", &["Region", "East", "West", "North"]).unwrap();
//! worksheet.write_column("C1", &["Sales Rep", "Tom", "Fred", "Amy"]).unwrap();
//! worksheet.write_column("C1", &["Product", "Apple", "Grape", "Pear"]).unwrap();
//! workbook.save_as("./examples/border_diagonal_cell.xlsx").unwrap();
//! ```
//!
//! Wrap merged cells
//! ```
//! use std::cell;
//! use edit_xlsx::{Cell, Format, FormatBorder, FormatBorderElement, FormatBorderType, FormatColor, Read, Workbook, Write};
//! let mut workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
//! let worksheet = workbook.get_worksheet_mut_by_name("worksheet").unwrap();
//! for row in 18..=21 {
//!     for col in 2..=11 {
//!         let mut cell: Cell<String> = worksheet.read_cell((row, col)).unwrap();
//!         cell.format = match cell.format.take() {
//!             None => None,
//!             Some(mut format) => Some(format.set_border(FormatBorderType::Double)),
//!         };
//!         worksheet.write_cell((row, col), &cell).unwrap()
//!     }
//! }
//! workbook.save_as("./examples/border_wrap_merged_cells.xlsx").unwrap();
//! ```
//!

use std::fmt::{Display, Formatter};
use crate::{Cell, FormatColor};
use crate::xml::common::FromFormat;
use crate::xml::style::border::{Border, BorderElement};
use crate::xml::style::color::Color;

///
/// [`FormatBorder`] is used to edit the border style of the [`Cell`] in each direction.
/// # Fields
/// | field        | type        | meaning                                                      |
/// | ------------ | ----------- | ------------------------------------------------------------ |
/// | `left`     | [`FormatBorderElement`] | The [`Cell`]'s left border style               |
/// | `right`    | [`FormatBorderElement`] | The [`Cell`]'s right border style               |
/// | `top`      | [`FormatBorderElement`] | The [`Cell`]'s top border style               |
/// | `bottom`   | [`FormatBorderElement`] | The [`Cell`]'s bottom border style               |
/// | `diagonal` | [`FormatBorderElement`] | The [`Cell`]'s diagonal border style               |
///
#[derive(Clone, Debug, PartialEq, Default)]
pub struct FormatBorder {
    pub left: FormatBorderElement,
    pub right: FormatBorderElement,
    pub top: FormatBorderElement,
    pub bottom: FormatBorderElement,
    pub diagonal: FormatBorderElement,
}

///
/// [`FormatBorderElement`] is used to edit one of the direction of the [`Cell`]'s border style.
/// # Fields
/// | field        | type        | meaning                                                      |
/// | ------------ | ----------- | ------------------------------------------------------------ |
/// | `border_type`     | [`FormatBorderType`] | Type of the border               |
/// | `color`    | [`FormatColor`] | Color of the border               |
///
#[derive(Clone, Debug, PartialEq, Copy, Default)]
pub struct FormatBorderElement {
    pub border_type: FormatBorderType,
    pub color: FormatColor,
}

impl FormatBorderElement {
    pub fn new(border_type: &FormatBorderType, color: &FormatColor) -> FormatBorderElement {
        FormatBorderElement {
            border_type: *border_type,
            color: *color,
        }
    }

    ///
    /// Create a new [`FormatBorderElement`] based on the border color,
    /// the border type defaults to [`FormatBorderType::Thin`]
    ///
    pub fn from_color(color: &FormatColor) -> FormatBorderElement {
        FormatBorderElement {
            border_type: FormatBorderType::Thin,
            color: *color,
        }
    }

    ///
    /// Create a new [`FormatBorderElement`] based on the border type,
    /// the border color defaults to [`FormatColor::Default`]
    ///
    pub fn from_border_type(border_type: &FormatBorderType) -> FormatBorderElement {
        FormatBorderElement {
            border_type: *border_type,
            color: FormatColor::default(),
        }
    }
}

///
/// Enumeration of different border type
///
/// # Fields:
/// | unit | meaning |
/// | ---- | ---- |
/// | `None` | Default, No border |
/// | `Thin` | Thin line |
/// | `Medium` | Medium line |
/// | `Dashed` | Dashed line |
/// | `Dotted` | Dotted line |
/// | `Thick` | Thick line |
/// | `Double` | Double line |
/// | `Hair` | Hairline |
/// | `MediumDashed` | Medium dashed line |
/// | `DashDot` | Dash-dot line |
/// | `MediumDashDot` | Medium dash-dot line |
/// | `DashDotDot` | Dash-dot-dot line |
/// | `MediumDashDotDot` | Medium dash-dot-dot line |
/// | `SlantDashDot` | Slanted dash-dot-dot line |
///
#[derive(Copy, Clone, Debug, PartialEq)]
pub enum FormatBorderType {
    /// Default, No border
    None,
    /// Thin line
    Thin,
    /// Medium line
    Medium,
    /// Dashed line
    Dashed,
    /// Dotted line
    Dotted,
    /// Thick line
    Thick,
    /// Double line
    Double,
    /// Hairline
    Hair,
    /// Medium dashed line
    MediumDashed,
    /// Dash-dot line
    DashDot,
    /// Medium dash-dot line
    MediumDashDot,
    /// Dash-dot-dot line
    DashDotDot,
    /// Medium dash-dot-dot line
    MediumDashDotDot,
    /// Slanted dash-dot-dot line
    SlantDashDot,
}

impl Display for FormatBorderType {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.write_str(self.to_str())
    }
}

impl Default for FormatBorderType {
    fn default() -> FormatBorderType {
        FormatBorderType::None
    }
}

impl FormatBorderType {
    pub(crate) fn to_str(&self) -> &str {
        match self {
            FormatBorderType::None => "none",
            FormatBorderType::Thin => "thin",
            FormatBorderType::Medium => "medium",
            FormatBorderType::Dashed => "dashed",
            FormatBorderType::Dotted => "dotted",
            FormatBorderType::Thick => "thick",
            FormatBorderType::Double => "double",
            FormatBorderType::Hair => "hair",
            FormatBorderType::MediumDashed => "mediumDashed",
            FormatBorderType::DashDot => "dashDot",
            FormatBorderType::MediumDashDot => "mediumDashDot",
            FormatBorderType::DashDotDot => "dashDotDot",
            FormatBorderType::MediumDashDotDot => "mediumDashDotDot",
            FormatBorderType::SlantDashDot => "slantDashDot",
        }
    }

    pub(crate) fn from_str(border_str: &str) -> Self {
        match border_str {
            "thin" => FormatBorderType::Thin,
            "medium" => FormatBorderType::Medium,
            "dashed" => FormatBorderType::Dashed,
            "dotted" => FormatBorderType::Dotted,
            "thick" => FormatBorderType::Thick,
            "double" => FormatBorderType::Double,
            "hair" => FormatBorderType::Hair,
            "mediumDashed" => FormatBorderType::MediumDashed,
            "dashDot" => FormatBorderType::DashDot,
            "mediumDashDot" => FormatBorderType::MediumDashDot,
            "dashDotDot" => FormatBorderType::DashDotDot,
            "mediumDashDotDot" => FormatBorderType::MediumDashDotDot,
            "slantDashDot" => FormatBorderType::SlantDashDot,
            _ => FormatBorderType::None,
        }
    }
}

impl FromFormat<FormatBorder> for Border {
    fn set_attrs_by_format(&mut self, format: &FormatBorder) {
        self.left = Some(BorderElement::from_format(&format.left));
        self.right = Some(BorderElement::from_format(&format.right));
        self.top = Some(BorderElement::from_format(&format.top));
        self.bottom = Some(BorderElement::from_format(&format.bottom));
        self.diagonal = Some(BorderElement::from_format(&format.diagonal));
    }

    fn set_format(&self, format: &mut FormatBorder) {
        format.left = {
            if let Some(left) = &self.left {
                left.get_format()
            } else {
                FormatBorderElement::default()
            }
        };
        format.right = {
            if let Some(right) = &self.right {
                right.get_format()
            } else {
                FormatBorderElement::default()
            }
        };
        format.top = {
            if let Some(top) = &self.top {
                top.get_format()
            } else {
                FormatBorderElement::default()
            }
        };
        format.bottom = {
            if let Some(bottom) = &self.bottom {
                bottom.get_format()
            } else {
                FormatBorderElement::default()
            }
        };
        format.diagonal = {
            if let Some(diagonal) = &self.diagonal {
                diagonal.get_format()
            } else {
                FormatBorderElement::default()
            }
        }
    }
}

impl FromFormat<FormatBorderElement> for BorderElement {
    fn set_attrs_by_format(&mut self, format: &FormatBorderElement) {
        self.style = Some(String::from(format.border_type.to_str()));
        self.color = Some(Color::from_format(&format.color));
    }

    fn set_format(&self, format: &mut FormatBorderElement) {
        // format.color = self.color.unwrap_or_default();
        match &self.style {
            None => format.border_type = FormatBorderType::default(),
            Some(style) => format.border_type = FormatBorderType::from_str(style)
        }
    }
}