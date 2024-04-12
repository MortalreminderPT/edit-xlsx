use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Deserialize, Serialize)]
pub(crate) struct DataValidations {
    #[serde(rename = "@count", default)]
    count: u32,
    #[serde(rename = "dataValidation", default)]
    data_validation: Vec<DataValidation>
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct DataValidation {
    #[serde(rename = "@type", default, skip_serializing_if = "Option::is_none")]
    tp: Option<String>,
    #[serde(rename = "@allowBlank", default, skip_serializing_if = "Option::is_none")]
    allow_blank: Option<u8>,
    #[serde(rename = "@showInputMessage", default, skip_serializing_if = "Option::is_none")]
    show_input_message: Option<u8>,
    #[serde(rename = "@showErrorMessage", default, skip_serializing_if = "Option::is_none")]
    show_error_message: Option<u8>,
    #[serde(rename = "@sqref", default, skip_serializing_if = "Option::is_none")]
    sqref: Option<String>,
    #[serde(rename = "formula1", default, skip_serializing_if = "Option::is_none")]
    formula1: Option<String>,
}