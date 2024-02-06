use serde::{Deserialize, Serialize};
use crate::xml::common::Color;


#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Borders {
    #[serde(rename = "@count", default)]
    count: u32,
    border: Vec<Border>,
}

impl Borders {
    pub(crate) fn default() -> Borders {
        Borders {
            count: 0,
            border: vec![],
        }
    }

    pub(crate) fn add_border(&mut self, border: &Border) -> u32 {
        for i in 0..self.border.len() {
            if self.border[i] == *border {
                return i as u32;
            }
        }
        self.count += 1;
        self.border.push(border.clone());
        self.border.len() as u32 - 1
    }
}

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
