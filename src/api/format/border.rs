use std::fmt::{Display, Formatter};
use crate::FormatColor;

#[derive(Clone)]
pub struct FormatBorder<'a> {
    pub left: FormatBorderElement<'a>,
    pub right: FormatBorderElement<'a>,
    pub top: FormatBorderElement<'a>,
    pub bottom: FormatBorderElement<'a>,
    pub diagonal: FormatBorderElement<'a>,
}

#[derive(Copy, Clone)]
pub struct FormatBorderElement<'a> {
    pub border_type: FormatBorderType,
    pub color: FormatColor<'a>,
}

impl Default for FormatBorderElement<'_> {
    fn default() -> Self {
        Self {
            border_type: Default::default(),
            color: Default::default(),
        }
    }
}

impl Default for FormatBorder<'_> {
    fn default() -> Self {
        Self {
            left: Default::default(),
            right: Default::default(),
            top: Default::default(),
            bottom: Default::default(),
            diagonal: Default::default(),
        }
    }
}

#[derive(Copy, Clone)]
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