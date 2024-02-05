use serde::{Deserialize, Serialize};
use crate::xml::common::Color;

#[derive(Debug, Deserialize, Serialize, Clone, PartialEq)]
pub(crate) struct Border {
    pub(crate) left: BorderElement,
    pub(crate) right: BorderElement,
    pub(crate) top: BorderElement,
    pub(crate) bottom: BorderElement,
    pub(crate) diagonal: BorderElement,
}

impl Border {
    pub(crate) fn default() -> Border {
        Border {
            left: BorderElement::default(),
            right: BorderElement::default(),
            top: BorderElement::default(),
            bottom: BorderElement::default(),
            diagonal: BorderElement::default(),
        }
    }
}

#[derive(Debug, Deserialize, Serialize, Clone, PartialEq)]
pub(crate) struct BorderElement {
    #[serde(rename = "@style", skip_serializing_if = "Option::is_none")]
    style: Option<String>,
    #[serde(rename = "color", skip_serializing_if = "Option::is_none")]
    color: Option<Color>,
}

impl BorderElement {
    pub(crate) fn new(style: &str, color: u32) -> BorderElement {
        BorderElement {
            style: Some(String::from(style)),
            color: Some(Color::default()),
        }
    }

    fn default() -> BorderElement {
        BorderElement {
            style: None,
            color: None,
        }
    }
}
