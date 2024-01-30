use serde::{Deserialize, Serialize};
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub(crate) struct PhoneticPr {
    #[serde(rename = "@fontId")]
    font_id: u32,
    #[serde(rename = "@type")]
    phonetic_pr_type: String
}