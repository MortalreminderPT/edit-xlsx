use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Border {
    left: BorderElement,
    right: BorderElement,
    top: BorderElement,
    bottom: BorderElement,
    diagonal: BorderElement,
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
struct BorderElement {
    #[serde(rename = "@style", skip_serializing_if = "Option::is_none")]
    style: Option<String>,
    #[serde(rename = "color", skip_serializing_if = "Option::is_none")]
    color: Option<BorderColor>,
}

impl BorderElement {
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
