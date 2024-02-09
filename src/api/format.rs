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

#[derive(Default)]
pub struct Format<'a> {
    pub(crate) font: FormatFont<'a>,
    pub(crate) border: FormatBorder<'a>,
    pub(crate) fill: FormatFill<'a>,
    pub(crate) align: FormatAlign,
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
        self.font.size = size;
        self
    }

    pub fn set_color<'b: 'a>(mut self, format_color: FormatColor<'b>) -> Self {
        self.font.color = format_color;
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
                self.align.horizontal = format_align_type,
            FormatAlignType::Top | FormatAlignType::VerticalCenter | FormatAlignType::Bottom =>
                self.align.vertical = format_align_type,
        }
        self
    }
}