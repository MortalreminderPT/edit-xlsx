use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub(crate) struct Color {
    #[serde(rename = "@indexed", skip_serializing_if = "Option::is_none")]
    pub(crate) indexed: Option<u8>,
    #[serde(rename = "@rgb", skip_serializing_if = "Option::is_none")]
    pub(crate) rgb: Option<String>,
    #[serde(rename = "@theme", skip_serializing_if = "Option::is_none")]
    pub(crate) theme: Option<u8>,
    #[serde(rename = "@tint", skip_serializing_if = "Option::is_none")]
    pub(crate) tint: Option<f64>,
    #[serde(rename = "@auto", skip_serializing_if = "Option::is_none")]
    auto: Option<u32>
}

impl Color {
    pub(crate) fn from_rgb(rgb: &str) -> Color {
        Color { indexed: None, rgb: Some(rgb.to_string()), theme: None, tint: None, auto: None }
    }

    pub(crate) fn from_index(id: u8) -> Color {
        Color { indexed: Some(id), rgb: None, theme: None, tint: None, auto: None }
    }

    pub(crate) fn from_theme(theme: u8, tint: f64) -> Color {
        Color { indexed: None, rgb: None, theme: Some(theme), tint: Some(tint), auto: None }
    }
}
