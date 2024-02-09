use serde::{Deserialize, Serialize};
use crate::api::format::FormatAlign;
use crate::xml::common::FromFormat;

#[derive(Debug, Deserialize, Serialize, PartialEq, Clone)]
pub(crate) struct Alignment {
    #[serde(rename = "@horizontal", skip_serializing_if = "Option::is_none")]
    pub(crate) horizontal: Option<String>,
    #[serde(rename = "@vertical", skip_serializing_if = "Option::is_none")]
    pub(crate) vertical: Option<String>,
}

impl Default for Alignment {
    fn default() -> Self {
        Alignment {
            horizontal: None,
            vertical: None,
        }
    }
}

impl FromFormat<FormatAlign> for Alignment {
    fn set_attrs_by_format(&mut self, format: &FormatAlign) {
        self.vertical = Some(String::from(format.vertical.to_str()));
        self.horizontal = Some(String::from(format.horizontal.to_str()));
    }
}

impl Alignment {
    pub(crate) fn new(horizontal: Option<String>, vertical: Option<String>) -> Alignment {
        Alignment {
            horizontal,
            vertical,
        }
    }
}