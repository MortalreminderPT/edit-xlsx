use serde::{Deserialize, Serialize};
use crate::xml::common::Element;
use crate::xml::style::Rearrange;

#[derive(Debug, Deserialize, Serialize, Clone, PartialEq)]
pub(crate) struct Font {
    sz: Element<u8>,
    #[serde(skip_serializing_if = "Option::is_none")]
    color: Option<Color>,
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

impl Font {
    pub(crate) fn default() -> Font {
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

#[derive(Debug, Deserialize, Serialize, Clone, PartialEq)]
struct Color {
    #[serde(rename = "@theme")]
    theme: u32
}
