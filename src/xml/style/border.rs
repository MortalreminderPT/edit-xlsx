use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize)]
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

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct BorderElement {
    #[serde(rename = "@style", skip_serializing_if = "Option::is_none")]
    style: Option<String>,
    #[serde(rename = "color", skip_serializing_if = "Option::is_none")]
    color: Option<BorderColor>,
}

impl BorderElement {
    pub(crate) fn new(style: &str, color: u32) -> BorderElement {
        BorderElement {
            style: Some(String::from(style)),
            color: Some(BorderColor::new(color)),
        }
    }

    fn default() -> BorderElement {
        BorderElement {
            style: None,
            color: None,
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct BorderColor {
    #[serde(rename = "@indexed")]
    indexed: u32,
}

impl BorderColor {
    fn new(indexed: u32) -> BorderColor {
        BorderColor {
            indexed
        }
    }
}