use crate::xml::common::FromFormat;

#[derive(Copy, Clone, Debug)]
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

#[derive(Clone, Debug)]
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