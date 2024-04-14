use serde::{Deserialize, Serialize};
use crate::xml::common;
use crate::xml::style::color::Color;

#[derive(Debug, Clone, Deserialize, Serialize)]
pub(crate) struct Borders {
    #[serde(rename = "@count", default, skip_serializing_if = "common::is_zero")]
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

    pub(crate) fn get_border(&self, id: u32) -> Option<&Border> {
        self.border.get(id as usize)
    }
}

#[derive(Debug, Deserialize, Serialize, Clone, PartialEq, Default)]
pub(crate) struct Border {
    #[serde(rename = "@diagonalUp", skip_serializing_if = "Option::is_none")]
    pub(crate) diagonal_up: Option<u8>,
    #[serde(rename = "@diagonalDown", skip_serializing_if = "Option::is_none")]
    pub(crate) diagonal_down: Option<u8>,
    #[serde(rename = "@outline", skip_serializing_if = "Option::is_none")]
    pub(crate) outline: Option<u8>,
    #[serde(rename = "left", skip_serializing_if = "Option::is_none")]
    pub(crate) left: Option<BorderElement>,
    #[serde(rename = "right", skip_serializing_if = "Option::is_none")]
    pub(crate) right: Option<BorderElement>,
    #[serde(rename = "top", skip_serializing_if = "Option::is_none")]
    pub(crate) top: Option<BorderElement>,
    #[serde(rename = "bottom", skip_serializing_if = "Option::is_none")]
    pub(crate) bottom: Option<BorderElement>,
    #[serde(rename = "diagonal", skip_serializing_if = "Option::is_none")]
    pub(crate) diagonal: Option<BorderElement>,
}

#[derive(Debug, Deserialize, Serialize, Clone, PartialEq)]
pub(crate) struct BorderElement {
    #[serde(rename = "@style", skip_serializing_if = "Option::is_none")]
    pub(crate) style: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) color: Option<Color>,
}

impl Default for BorderElement {
    fn default() -> Self {
        BorderElement {
            style: None,
            color: Some(Color::default()),
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
