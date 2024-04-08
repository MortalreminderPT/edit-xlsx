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
#[derive(Clone)]
pub struct FormatAlign {
    pub(crate) horizontal: Option<FormatAlignType>,
    pub(crate) vertical: Option<FormatAlignType>,
    pub(crate) reading_order: Option<u8>,
    pub(crate) indent: Option<u8>,
}

impl Default for FormatAlign {
    fn default() -> Self {
        Self {
            horizontal: None,// FormatAlignType::Right,
            vertical: None,// FormatAlignType::Bottom,
            reading_order: None,
            indent: None,
        }
    }
}