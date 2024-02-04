use crate::xml::style::{Bold, Font, Italic, Underline};

pub struct Format {
    pub(crate) font: Option<Font>,
}

impl Format {
    pub fn new() -> Format {
        Format {
            font: None
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
}