use std::fmt::Debug;
pub use crate::api::align::FormatAlign;
pub use crate::api::border::FormatBorder;
pub use crate::api::color::FormatColor;
use crate::xml::style::border::{Border, BorderElement, FromFormat};
use crate::xml::style::fill::Fill;
use crate::xml::style::xf::Xf;
use crate::xml::common::Element;
use crate::xml::style::alignment::Alignment;
use crate::xml::style::color::Color;
use crate::xml::style::font::{Bold, Font, Italic, Underline};

pub struct Format {
    pub(crate) font: Option<Font>,
    pub(crate) border: Option<Border>,
    pub(crate) fill: Option<Fill>,
    pub(crate) xf: Option<Xf>,
}

impl Format {
    pub fn new() -> Format {
        Format {
            font: None,
            border: None,
            fill: None,
            xf: None,
        }
    }

    pub fn set_bold(mut self) -> Format {
        let font = self.font.get_or_insert(Font::default());
        font.bold = Some(Bold::default());
        self
    }
    pub fn set_italic(mut self) -> Format {
        let font = self.font.get_or_insert(Font::default());
        font.italic = Some(Italic::default());
        self
    }
    pub fn set_underline(mut self) -> Format {
        let font = self.font.get_or_insert(Font::default());
        font.underline = Some(Underline::default());
        self
    }

    pub fn set_size(mut self, size: u8) -> Format {
        let font = self.font.get_or_insert(Font::default());
        font.sz = Element::from_format(&size); //from_val(size);
        self
    }

    pub fn set_color(mut self, color: FormatColor) -> Format {
        let font = self.font.get_or_insert(Font::default());
        font.color = Some(Color::from_format(&color));
        self
    }

    pub fn set_border(mut self, format_border: FormatBorder) -> Format {
        let mut border = self.border.get_or_insert(Border::default());
        *border = Border::from_format(&format_border);
        self
    }

    pub fn set_border_left(mut self, format_border: FormatBorder) -> Format {
        let border = self.border.get_or_insert(Border::default());
        border.left = BorderElement::from_format(&format_border);
        self
    }

    pub fn set_border_right(mut self, format_border: FormatBorder) -> Format {
        let border = self.border.get_or_insert(Border::default());
        border.right = BorderElement::from_format(&format_border);
        self
    }

    pub fn set_border_top(mut self, format_border: FormatBorder) -> Format {
        let border = self.border.get_or_insert(Border::default());
        border.top = BorderElement::from_format(&format_border);
        self
    }
    pub fn set_border_bottom(mut self, format_border: FormatBorder) -> Format {
        let border = self.border.get_or_insert(Border::default());
        border.bottom = BorderElement::from_format(&format_border);
        self
    }
    pub fn set_background_color(mut self, color: FormatColor) -> Format {
        let fill = self.fill.get_or_insert(Fill::default());
        fill.pattern_fill.pattern_type = "solid".to_string();
        fill.pattern_fill.fg_color = Some(Color::from_format(&color));
        self
    }

    pub fn set_align(mut self, format_align: FormatAlign) -> Format {
        let xf = self.xf.get_or_insert(Xf::default());
        let align = xf.alignment.get_or_insert(Alignment::default());
        match format_align {
            FormatAlign::Top => {
                align.vertical = Some(String::from("top"));
            }
            FormatAlign::Center => {
                align.vertical = Some(String::from("center"));
            }
            FormatAlign::Bottom => {
                align.vertical = Some(String::from("bottom"));
            }
            FormatAlign::Left => {
                align.horizontal = Some(String::from("left"));
            }
            FormatAlign::VerticalCenter => {
                align.horizontal = Some(String::from("center"));
            }
            FormatAlign::Right => {
                align.horizontal = Some(String::from("right"));
            }
        }
        self
    }
}