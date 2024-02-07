pub enum FormatBorder {
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


impl FormatBorder {
    pub(crate) fn to_str(&self) -> &str {
        match self {
            FormatBorder::None => "none",
            FormatBorder::Thin => "thin",
            FormatBorder::Medium => "medium",
            FormatBorder::Dashed => "dashed",
            FormatBorder::Dotted => "dotted",
            FormatBorder::Thick => "thick",
            FormatBorder::Double => "double",
            FormatBorder::Hair => "hair",
            FormatBorder::MediumDashed => "mediumDashed",
            FormatBorder::DashDot => "dashDot",
            FormatBorder::MediumDashDot => "mediumDashDot",
            FormatBorder::DashDotDot => "dashDotDot",
            FormatBorder::MediumDashDotDot => "mediumDashDotDot",
            FormatBorder::SlantDashDot => "slantDashDot",
        }
    }
}