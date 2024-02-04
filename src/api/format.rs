use crate::xml::style::border::Border;
use crate::xml::style::font::{Bold, Font, Italic, Underline};

pub struct Format {
    pub(crate) font: Option<Font>,
    pub(crate) border: Option<Border>,
}

impl Format {
    pub fn new() -> Format {
        Format {
            font: None,
            border: None
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

    pub fn set_border(mut self) -> Format {
        let border = self.border.get_or_insert(Border::default());
        self
    }
}