use serde::{Deserialize, Serialize};
use crate::xml::common::{Element, FromFormat};
use crate::xml::style::color::Color;
use crate::api::format::FormatFont;

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Fonts {
    #[serde(rename = "@count", default)]
    count: u32,
    #[serde(rename(serialize = "@x14ac:knownFonts", deserialize = "@knownFonts"), default)]
    x14ac_known_fonts: u32,
    #[serde(rename = "font", default)]
    fonts: Vec<Font>,
}

impl Default for Fonts {
    fn default() -> Self {
        Fonts {
            count: 1,
            x14ac_known_fonts: 1,
            fonts: vec![Default::default()],
        }
    }
}

impl Fonts {
    pub(crate) fn add_font(&mut self, font: &Font) -> u32 {
        for i in 0..self.fonts.len() {
            if self.fonts[i] == *font {
                return i as u32;
            }
        }
        self.count += 1;
        self.fonts.push(font.clone());
        self.fonts.len() as u32 - 1
    }
}

#[derive(Debug, Deserialize, Serialize, Clone, PartialEq)]
pub(crate) struct Font {
    pub(crate) sz: Element<u8>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) color: Option<Color>,
    name: Element<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    family: Option<Element<u8>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    charset: Option<Element<u8>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    scheme: Option<Element<String>>,
    #[serde(rename = "b", skip_serializing_if = "Option::is_none")]
    pub(crate) bold: Option<Bold>,
    #[serde(rename = "i", skip_serializing_if = "Option::is_none")]
    pub(crate) italic: Option<Italic>,
    #[serde(rename = "u", skip_serializing_if = "Option::is_none")]
    pub(crate) underline: Option<Underline>,
}

impl Default for Font {
    fn default() -> Font {
        Font {
            sz: Element::from_val(11),
            color: None,
            name: Element::from_val("Calibri".to_string()),
            family: None,
            charset: None,
            scheme: None,
            bold: None,
            italic: None,
            underline: None,
        }
    }
}

impl FromFormat<FormatFont<'_>> for Font {
    fn set_attrs_by_format(&mut self, format: &FormatFont) {
        self.color = Some(Color::from_format(&format.color));
        self.sz = Element::from_val(format.size);
        self.bold = if format.bold { Some(Bold::default()) } else { None };
        self.underline = if format.underline { Some(Underline::default()) } else { None };
        self.italic = if format.italic { Some(Italic::default()) } else { None };
    }
}

#[derive(Debug, Deserialize, Serialize, Clone, PartialEq)]
pub(crate) struct Bold {

}

impl Bold {
    pub(crate) fn default() -> Bold {
        Bold {}
    }
}

#[derive(Debug, Deserialize, Serialize, Clone, PartialEq)]
pub(crate) struct Italic {

}

impl Italic {
    pub(crate) fn default() -> Italic {
        Italic {}
    }
}

#[derive(Debug, Deserialize, Serialize, Clone, PartialEq)]
pub(crate) struct Underline {

}

impl Underline {
    pub(crate) fn default() -> Underline {
        Underline {}
    }
}

// #[derive(Debug, Deserialize, Serialize, Clone, PartialEq)]
// struct Color {
//     #[serde(rename = "@theme")]
//     theme: u32
// }
