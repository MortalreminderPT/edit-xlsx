use std::fmt::Debug;
pub use crate::api::alignment::FormatAlign;
pub use crate::api::color::Color;
use crate::xml::style::border::{Border, BorderElement};
use crate::xml::style::fill::Fill;
use crate::xml::style::xf::Xf;
use crate::xml::common::{Color as XmlColor, Element};
use crate::xml::style::alignment::Alignment;
use crate::xml::style::font::{Bold, Font, Italic, Underline};

pub struct Format {
    pub(crate) font: Option<Font>,
    pub(crate) border: Option<Border>,
    pub(crate) fill: Option<Fill>,
    pub(crate) xf: Option<Xf>,
}

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
    fn to_str(&self) -> &str {
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
        font.sz = Element::from_val(size);
        self
    }

    pub fn set_color(mut self, color: Color) -> Format {
        match color {
            Color::RGB(rgb) => {
                let font = self.font.get_or_insert(Font::default());
                font.color = Some(XmlColor::from_rgb(rgb));
                self
            },
        }
    }

    pub fn set_border(mut self, format_border: FormatBorder) -> Format {
        let border = self.border.get_or_insert(Border::default());
        let style = format_border.to_str();
        border.left = BorderElement::new(style, 64);
        border.right = BorderElement::new(style, 64);
        border.top = BorderElement::new(style, 64);
        border.bottom = BorderElement::new(style, 64);
        self
    }
    pub fn set_border_left(mut self, format_border: FormatBorder) -> Format {
        let border = self.border.get_or_insert(Border::default());
        let style = format_border.to_str();
        border.left = BorderElement::new(style, 64);
        self
    }
    pub fn set_border_right(mut self, format_border: FormatBorder) -> Format {
        let border = self.border.get_or_insert(Border::default());
        let style = format_border.to_str();
        border.right = BorderElement::new(style, 64);
        self
    }
    pub fn set_border_top(mut self, format_border: FormatBorder) -> Format {
        let border = self.border.get_or_insert(Border::default());
        let style = format_border.to_str();
        border.top = BorderElement::new(style, 64);
        self
    }
    pub fn set_border_bottom(mut self, format_border: FormatBorder) -> Format {
        let border = self.border.get_or_insert(Border::default());
        let style = format_border.to_str();
        border.bottom = BorderElement::new(style, 64);
        self
    }
    pub fn set_background_color(mut self, color: Color) -> Format {
        match color {
            Color::RGB(rgb) => {
                let fill = self.fill.get_or_insert(Fill::default());
                fill.pattern_fill.pattern_type = "solid".to_string();
                fill.pattern_fill.fg_color = Some(XmlColor::from_rgb(rgb));
                self
            },
        }
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