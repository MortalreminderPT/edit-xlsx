use serde::{Deserialize, Serialize};
use crate::xml::style::font::Font;

#[derive(Debug, Clone, Deserialize, Serialize)]
pub(crate) struct InlineString {
    #[serde(rename = "r", skip_serializing_if = "Vec::is_empty")]
    rich_texts: Vec<RichText>
}

#[derive(Debug, Clone, Deserialize, Serialize)]
pub(crate) struct RichText {
    #[serde(rename = "rPr", skip_serializing_if = "Option::is_none")]
    font: Option<Font>,
    #[serde(rename = "t", skip_serializing_if = "String::is_empty")]
    text: String,
}