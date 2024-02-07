use serde::{Deserialize, Serialize};
use crate::api::border::FormatBorder;
pub(crate) use crate::xml::common::FromFormat;
use crate::xml::style::color::Color;

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Borders {
    #[serde(rename = "@count", default)]
    count: u32,
    border: Vec<Border>,
}

impl Borders {
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

#[derive(Debug, Deserialize, Serialize, Clone, PartialEq)]
pub(crate) struct BorderElement {
    #[serde(rename = "@style", skip_serializing_if = "Option::is_none")]
    style: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    color: Option<Color>,
}

impl BorderElement {
    pub(crate) fn new(style: &str, color: u32) -> BorderElement {
        BorderElement {
            style: Some(String::from(style)),
            color: Some(Color::default()),
        }
    }
}

impl Default for BorderElement {
    fn default() -> Self {
        BorderElement {
            style: None,
            color: Some(Color::default()),
        }
    }
}

impl Default for Border {
    fn default() -> Self {
        Border {
            left: BorderElement::default(),
            right: BorderElement::default(),
            top: BorderElement::default(),
            bottom: BorderElement::default(),
            diagonal: BorderElement::default(),
        }
    }
}

impl Default for Borders {
    fn default() -> Self {
        Borders {
            count: 0,
            border: vec![],
        }
    }
}

impl FromFormat<FormatBorder> for Border {
    fn set_attrs_by_format(&mut self, format: &FormatBorder) {
        let style = format.to_str();
        self.left = BorderElement::new(style, 64);
        self.right = BorderElement::new(style, 64);
        self.top = BorderElement::new(style, 64);
        self.bottom = BorderElement::new(style, 64);
    }
}

impl FromFormat<FormatBorder> for BorderElement {
    fn set_attrs_by_format(&mut self, format: &FormatBorder) {
        let style = format.to_str();
        *self = BorderElement::new(style, 64);
    }
}