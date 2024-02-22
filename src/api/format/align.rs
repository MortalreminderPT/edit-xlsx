#[derive(Copy, Clone)]
pub enum FormatAlignType {
    Top,
    Center,
    Bottom,
    Left,
    VerticalCenter,
    Right,
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
}

pub struct FormatAlign {
    pub(crate) horizontal: FormatAlignType,
    pub(crate) vertical: FormatAlignType,
    pub(crate) reading_order: Option<u8>,
}

impl Default for FormatAlign {
    fn default() -> Self {
        Self {
            horizontal: FormatAlignType::Right,
            vertical: FormatAlignType::Bottom,
            reading_order: None,
        }
    }
}