pub use
crate::api::format::align::FormatAlign;
pub use crate::api::format::align::FormatAlignType;
pub use crate::api::format::border::{FormatBorder, FormatBorderElement, FormatBorderType};
pub use color::FormatColor;
pub use crate::api::format::fill::FormatFill;
pub use font::FormatFont;

mod align;
mod color;
mod fill;
mod font;
pub mod border;

#[derive(Default, Clone)]
pub struct Format<'a> {
    pub(crate) font: FormatFont<'a>,
    pub(crate) border: FormatBorder<'a>,
    pub(crate) fill: FormatFill<'a>,
    pub(crate) align: FormatAlign,
}

impl<'a> Format<'a> {
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

impl<'a> Format<'a> {
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

    pub fn set_color<'b: 'a>(mut self, format_color: FormatColor<'b>) -> Self {
        self.font.color = format_color;
        self
    }

    pub fn set_font(mut self, font_name: &'a str) -> Self {
        self.font.name = font_name;
        self
    }

    pub fn set_border<'b: 'a>(mut self, format_border_type: FormatBorderType) -> Self {
        let mut format_border = FormatBorderElement::default();
        format_border.border_type = format_border_type;
        self.border.left = format_border;
        self.border.right = format_border;
        self.border.top = format_border;
        self.border.bottom = format_border;
        self.border.diagonal = format_border;
        self
    }

    pub fn set_border_left<'b: 'a>(mut self, format_border_type: FormatBorderType) -> Self {
        let mut format_border = FormatBorderElement::default();
        format_border.border_type = format_border_type;
        self.border.left = format_border;
        self
    }

    pub fn set_border_right<'b: 'a>(mut self, format_border_type: FormatBorderType) -> Self {
        let mut format_border = FormatBorderElement::default();
        format_border.border_type = format_border_type;
        self.border.right = format_border;
        self
    }

    pub fn set_border_top<'b: 'a>(mut self, format_border_type: FormatBorderType) -> Self {
        let mut format_border = FormatBorderElement::default();
        format_border.border_type = format_border_type;
        self.border.top = format_border;
        self
    }

    pub fn set_border_bottom<'b: 'a>(mut self, format_border_type: FormatBorderType) -> Self {
        let mut format_border = FormatBorderElement::default();
        format_border.border_type = format_border_type;
        self.border.bottom = format_border;
        self
    }

    pub fn set_background_color<'b: 'a>(mut self, format_color: FormatColor<'b>) -> Self {
        self.fill.pattern_type = "solid";
        self.fill.fg_color = format_color;
        self
    }

    pub fn set_align<'b: 'a>(mut self, format_align_type: FormatAlignType) -> Self {
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