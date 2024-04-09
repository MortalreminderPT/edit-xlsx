pub use
crate::api::format::align::FormatAlign;
pub use crate::api::format::align::FormatAlignType;
pub use crate::api::format::border::{FormatBorder, FormatBorderElement, FormatBorderType};
pub use color::FormatColor;
pub use crate::api::format::fill::FormatFill;
pub use font::FormatFont;
use crate::api::cell::values::CellType::String;

mod align;
mod color;
mod fill;
mod font;
pub mod border;

#[derive(Default, Clone)]
pub struct Format {
    pub(crate) font: FormatFont,
    pub(crate) border: FormatBorder,
    pub(crate) fill: FormatFill,
    pub(crate) align: FormatAlign,
}

impl Format {
    pub fn is_bold(&self) -> bool {
        self.font.bold
    }

    pub fn is_italic(&self) -> bool {
        self.font.italic
    }

    pub fn is_underline(&self) -> bool {
        self.font.underline
    }

    pub fn get_size(&self) -> f64 {
        self.font.size
    }

    pub fn get_borders(&self) -> &FormatBorder {
        &self.border
    }
}

impl Format {
    pub fn set_bold(mut self) -> Self {
        self.font.bold = true;
        self
    }

    pub fn set_italic(mut self) -> Self {
        self.font.italic = true;
        self
    }

    pub fn set_underline(mut self) -> Self {
        self.font.underline = true;
        self
    }

    pub fn set_size(mut self, size: u8) -> Self {
        self.font.size = size as f64;
        self
    }

    pub fn set_size_f64(mut self, size: f64) -> Self {
        self.font.size = size;
        self
    }

    pub fn set_color(mut self, format_color: FormatColor) -> Self {
        self.font.color = format_color;
        self
    }

    pub fn set_font(mut self, font_name: &str) -> Self {
        self.font.name = font_name.to_string();
        self
    }

    pub fn set_border(mut self, format_border_type: FormatBorderType) -> Self {
        let mut format_border = FormatBorderElement::default();
        format_border.border_type = format_border_type;
        self.border.left = format_border.clone();
        self.border.right = format_border.clone();
        self.border.top = format_border.clone();
        self.border.bottom = format_border.clone();
        self.border.diagonal = format_border;
        self
    }

    pub fn set_border_left(mut self, format_border_type: FormatBorderType) -> Self {
        let mut format_border = FormatBorderElement::default();
        format_border.border_type = format_border_type;
        self.border.left = format_border;
        self
    }

    pub fn set_border_right(mut self, format_border_type: FormatBorderType) -> Self {
        let mut format_border = FormatBorderElement::default();
        format_border.border_type = format_border_type;
        self.border.right = format_border;
        self
    }

    pub fn set_border_top(mut self, format_border_type: FormatBorderType) -> Self {
        let mut format_border = FormatBorderElement::default();
        format_border.border_type = format_border_type;
        self.border.top = format_border;
        self
    }

    pub fn set_border_bottom(mut self, format_border_type: FormatBorderType) -> Self {
        let mut format_border = FormatBorderElement::default();
        format_border.border_type = format_border_type;
        self.border.bottom = format_border;
        self
    }

    pub fn set_background_color(mut self, format_color: FormatColor) -> Self {
        self.fill.pattern_type = "solid".to_string();
        self.fill.fg_color = format_color;
        self
    }

    pub fn set_align(mut self, format_align_type: FormatAlignType) -> Self {
        match format_align_type {
            FormatAlignType::Left | FormatAlignType::Center | FormatAlignType::Right =>
                self.align.horizontal = Some(format_align_type),
            FormatAlignType::Top | FormatAlignType::VerticalCenter | FormatAlignType::Bottom =>
                self.align.vertical = Some(format_align_type),
        }
        self
    }

    pub fn set_reading_order(mut self, reading_order: u8) -> Self {
        self.align.reading_order = Some(reading_order);
        self
    }

    pub fn set_indent(mut self, indent: u8) -> Self {
        self.align.indent = Some(indent);
        self
    }
}