use serde::{Deserialize, Serialize};
use crate::api::format::FormatAlign;
use crate::xml::common::FromFormat;

#[derive(Debug, Deserialize, Serialize, PartialEq, Clone)]
pub(crate) struct Alignment {
    #[serde(rename = "@horizontal", skip_serializing_if = "Option::is_none")]
    pub(crate) horizontal: Option<String>,
    #[serde(rename = "@vertical", skip_serializing_if = "Option::is_none")]
    pub(crate) vertical: Option<String>,
    #[serde(rename = "@readingOrder", skip_serializing_if = "Option::is_none")]
    reading_order: Option<u8>,
    #[serde(rename = "@wrapText", skip_serializing_if = "Option::is_none")]
    wrap_text: Option<u8>,
    #[serde(rename = "@indent", skip_serializing_if = "Option::is_none")]
    indent: Option<u8>,
}

impl Default for Alignment {
    fn default() -> Self {
        Alignment {
            horizontal: None,
            vertical: None,
            reading_order: None,
            wrap_text: None,
            indent: None,
        }
    }
}

impl FromFormat<FormatAlign> for Alignment {
    fn set_attrs_by_format(&mut self, format: &FormatAlign) {
        if let Some(vertical) = format.vertical {
            self.vertical = Some(String::from(vertical.to_str()));
        }
        if let Some(horizontal) = format.horizontal {
            self.horizontal = Some(String::from(horizontal.to_str()));
        }
        self.reading_order = format.reading_order;
        self.indent = format.indent;
    }
}

// impl Alignment {
//     pub(crate) fn new(horizontal: Option<String>, vertical: Option<String>) -> Alignment {
//         Alignment {
//             horizontal,
//             vertical,
//         }
//     }
// }