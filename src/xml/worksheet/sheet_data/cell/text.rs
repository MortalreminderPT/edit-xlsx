use serde::{Deserialize, Deserializer, Serialize};
use serde::de::IntoDeserializer;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub(crate) struct Text {
    #[serde(rename = "$text", skip_serializing_if = "String::is_empty")]
    pub(crate) text: String,
    #[serde(rename = "@xml:space", default, skip_serializing_if = "String::is_empty")]
    pub(crate) xml_space: String,
}

impl Text {
    pub(crate) fn new(text: &str) -> Text {
        Text {
            text: text.to_string(),
            xml_space: "".to_string(),
        }
    }

    pub(crate) fn new_with_space(text: &str) -> Text {
        Text {
            text: text.to_string(),
            xml_space: "preserve".to_string(),
        }
    }
}