use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Deserialize, Serialize)]
pub(crate) struct ConditionalFormatting {
    #[serde(rename = "cfRule", default, skip_serializing_if = "Vec::is_empty")]
    cf_rule: Vec<CfRule>,
    #[serde(rename = "@sqref", default, skip_serializing_if = "String::is_empty")]
    sqref: String,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct CfRule {
    #[serde(rename = "@type", default, skip_serializing_if = "Option::is_none")]
    tp: Option<String>,
    #[serde(rename = "@dxfId", default, skip_serializing_if = "Option::is_none")]
    dxf_id: Option<u8>,
    #[serde(rename = "@priority", default, skip_serializing_if = "Option::is_none")]
    priority: Option<u8>,
    #[serde(rename = "@operator", default, skip_serializing_if = "Option::is_none")]
    operator: Option<String>,
    #[serde(rename = "formula", default, skip_serializing_if = "Option::is_none")]
    formula: Option<String>,
}