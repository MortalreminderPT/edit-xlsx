use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize, Default)]
pub(crate) struct X14SlicerStyles {
    #[serde(rename = "@defaultSlicerStyle", skip_serializing_if = "String::is_empty")]
    default_slicer_style: String
}