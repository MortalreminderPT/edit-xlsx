use serde::{Deserialize, Serialize};
use crate::xml::common::{Element, FromFormat};
use crate::xml::style::color::Color;
use crate::api::format::FormatFont;
use crate::RichText;

#[derive(Debug, Clone, Deserialize, Serialize)]
pub(crate) struct Fonts {
    #[serde(rename = "@count", default)]
    count: u32,
    #[serde(rename(serialize = "@x14ac:knownFonts", deserialize = "@knownFonts"), skip_serializing_if = "Option::is_none")]
    x14ac_known_fonts: Option<u32>,
    #[serde(rename = "font", default)]
    fonts: Vec<Font>,
}

impl Default for Fonts {
    fn default() -> Self {
        Fonts {
            count: 1,
            x14ac_known_fonts: Some(1),
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
    
    pub(crate) fn get_font(&self, id: u32) -> Option<&Font> {
        self.fonts.get(id as usize)
    }
}

#[derive(Debug, Deserialize, Serialize, Clone, PartialEq)]
pub(crate) struct Font {
    #[serde(rename = "b", skip_serializing_if = "Option::is_none")]
    pub(crate) bold: Option<Bold>,
    #[serde(rename = "i", skip_serializing_if = "Option::is_none")]
    pub(crate) italic: Option<Italic>,
    #[serde(skip_serializing_if = "Option::is_none")]
    strike: Option<Element<u8>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    condense: Option<Element<u8>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    extend: Option<Element<u8>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    outline: Option<Element<u8>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    shadow: Option<Element<u8>>,
    #[serde(rename = "u", skip_serializing_if = "Option::is_none")]
    pub(crate) underline: Option<Underline>,
    #[serde(rename = "vertAlign", skip_serializing_if = "Option::is_none")]
    vert_align: Option<Element<String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) sz: Option<Element<f64>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) color: Option<Color>,
    #[serde(rename = "name", default, skip_serializing_if = "Option::is_none")]
    pub(crate) name: Option<Element<String>>,
    #[serde(rename = "rFont", default, skip_serializing_if = "Option::is_none")]
    pub(crate) rich_font_name: Option<Element<String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    family: Option<Element<u8>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    charset: Option<Element<u8>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    scheme: Option<Element<String>>,
}

impl Default for Font {
    fn default() -> Font {
        Font {
            sz: Some(Element::from_val(11.0)),
            color: None,
            name: Some(Element::from_val("Calibri".to_string())),
            rich_font_name: None,
            family: None,
            charset: None,
            scheme: None,
            bold: None,
            italic: None,
            strike: None,
            condense: None,
            extend: None,
            outline: None,
            shadow: None,
            underline: None,
            vert_align: None,
        }
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

impl Font {
    pub(crate) fn from_rich_font_format(rich_font: &FormatFont) -> Font {
        let mut font = Font::from_format(rich_font);
        font.rich_font_name = font.name;
        font.name = None;
        font
    }

    pub(crate) fn get_rich_font_format(&self) -> FormatFont {
        let mut rich_font = self.get_format();
        if let Some(rich_font_name) = &self.rich_font_name {
            rich_font.name = rich_font_name.val.to_string();
        }
        rich_font
    }
}