use crate::FormatColor;
use crate::xml::common::{Element, FromFormat};
use crate::xml::style::color::Color;
use crate::xml::style::font::{Bold, Font, Italic, Underline};

#[derive(Clone, Debug)]
pub struct FormatFont {
    pub(crate) bold: bool,
    pub(crate) italic: bool,
    pub(crate) underline: bool,
    pub(crate) size: f64,
    pub(crate) color: FormatColor,
    pub(crate) name: String,// &'a str,
}

impl Default for FormatFont {
    fn default() -> Self {
        FormatFont {
            bold: false,
            italic: false,
            underline: false,
            size: 11.0,
            color: Default::default(),
            name: "Calibri".to_string(),
        }
    }
}

impl FromFormat<FormatFont> for Font {
    fn set_attrs_by_format(&mut self, format: &FormatFont) {
        self.color = Some(Color::from_format(&format.color));
        self.name = Some(Element::from_val(format.name.to_string()));
        self.sz = Some(Element::from_val(format.size));
        self.bold = if format.bold { Some(Bold::default()) } else { None };
        self.underline = if format.underline { Some(Underline::default()) } else { None };
        self.italic = if format.italic { Some(Italic::default()) } else { None };
    }

    fn set_format(&self, format: &mut FormatFont) {
        format.bold = self.bold.is_some();
        format.italic = self.italic.is_some();
        format.underline = self.underline.is_some();
        if let Some(size) = &self.sz {
            format.size = size.get_format();
        }
        if let Some(name) = &self.name {
            format.name = name.val.to_string();
        }
        format.color = self.color.as_ref().get_format();
    }
}
