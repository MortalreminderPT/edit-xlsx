use serde::{Deserialize, Serialize};
use crate::FormatColor;
use crate::xml::common::FromFormat;

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub(crate) struct Color {
    #[serde(rename = "@indexed", skip_serializing_if = "Option::is_none")]
    indexed: Option<u32>,
    #[serde(rename = "@rgb", skip_serializing_if = "Option::is_none")]
    rgb: Option<String>,
    #[serde(rename = "@theme", skip_serializing_if = "Option::is_none")]
    theme: Option<u32>,
    #[serde(rename = "@tint", skip_serializing_if = "Option::is_none")]
    tint: Option<f64>,
    #[serde(rename = "@auto", skip_serializing_if = "Option::is_none")]
    auto: Option<u32>
}

impl Color {
    pub(crate) fn from_rgb(rgb: &str) -> Color {
        Color { indexed: None, rgb: Some(rgb.to_string()), theme: None, tint: None, auto: None }
    }
}

impl FromFormat<FormatColor<'_>> for Color {
    fn set_attrs_by_format(&mut self, format: &FormatColor) {
        match format {
            FormatColor::Default => return,
            FormatColor::RGB(color) => {
                *self = Color::from_rgb(color)
            }
        }
    }
}