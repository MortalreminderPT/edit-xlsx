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
    pub(crate) fn from_rgb(r: u8, g: u8, b: u8) -> Color {
        let rgb = format!("FF{:02X}{:02X}{:02X}", r, g, b);
        Color { indexed: None, rgb: Some(rgb), theme: None, tint: None, auto: None }
    }

    pub(crate) fn from_index(id: u8) -> Color {
        Color { indexed: Some(id), rgb: None, theme: None, tint: None, auto: None }
    }

    pub(crate) fn from_theme(theme: u8, tint: f64) -> Color {
        Color { indexed: None, rgb: None, theme: Some(theme), tint: Some(tint), auto: None }
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.theme.is_none()
            && self.tint.is_none()
            && self.rgb.is_none()
            && self.indexed.is_none()
            && self.auto.is_none()
    }
}
