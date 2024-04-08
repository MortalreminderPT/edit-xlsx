use crate::FormatColor;

#[derive(Clone)]
pub struct FormatBorder<'a> {
    pub(crate) left: FormatBorderElement<'a>,
    pub(crate) right: FormatBorderElement<'a>,
    pub(crate) top: FormatBorderElement<'a>,
    pub(crate) bottom: FormatBorderElement<'a>,
    pub(crate) diagonal: FormatBorderElement<'a>,
}

#[derive(Copy, Clone)]
pub struct FormatBorderElement<'a> {
    pub(crate) border_type: FormatBorderType,
    pub(crate) color: FormatColor<'a>,
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
}