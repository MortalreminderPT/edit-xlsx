use crate::FormatColor;
use crate::xml::style::font::Font;

#[derive(Clone)]
pub struct FormatFont<'a> {
    pub(crate) bold: bool,
    pub(crate) italic: bool,
    pub(crate) underline: bool,
    pub(crate) size: f64,
    pub(crate) color: FormatColor<'a>,
    pub(crate) name: &'a str,
}

impl Default for FormatFont<'_> {
    fn default() -> Self {
        FormatFont {
            bold: false,
            italic: false,
            underline: false,
            size: 11.0,
            color: Default::default(),
            name: "Calibri",
        }
    }
}

impl FormatFont<'_> {
    pub(crate) fn from_font(font: &Font) -> Self {
        let mut format_font = FormatFont::default();
        format_font.bold = font.bold.is_some();
        format_font.italic = font.italic.is_some();
        format_font.underline = font.underline.is_some();
        format_font.size = font.sz.val;
        format_font
    }
}