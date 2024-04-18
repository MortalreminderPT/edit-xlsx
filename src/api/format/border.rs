use std::fmt::{Display, Formatter};
use crate::FormatColor;
use crate::xml::common::FromFormat;
use crate::xml::style::border::{Border, BorderElement};
use crate::xml::style::color::Color;

#[derive(Clone, Debug, PartialEq, Default)]
pub struct FormatBorder {
    pub left: FormatBorderElement,
    pub right: FormatBorderElement,
    pub top: FormatBorderElement,
    pub bottom: FormatBorderElement,
    pub diagonal: FormatBorderElement,
}

#[derive(Clone, Debug, PartialEq, Default)]
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

    pub fn from_color(color: &FormatColor) -> FormatBorderElement {
        FormatBorderElement {
            border_type: FormatBorderType::Thin,
            color: *color,
        }
    }

    pub fn from_border_type(border_type: &FormatBorderType) -> FormatBorderElement {
        FormatBorderElement {
            border_type: *border_type,
            color: FormatColor::default(),
        }
    }
}

#[derive(Copy, Clone, Debug, PartialEq)]
pub enum FormatBorderType {
    None,
    Thin,
    Medium,
    Dashed,
    Dotted,
    Thick,
    Double,
    Hair,
    MediumDashed,
    DashDot,
    MediumDashDot,
    DashDotDot,
    MediumDashDotDot,
    SlantDashDot,
}

impl Display for FormatBorderType {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.write_str(self.to_str())
    }
}

impl Default for FormatBorderType {
    fn default() -> Self {
        Self::None
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