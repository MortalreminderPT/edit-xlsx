use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct NumFmts {
    #[serde(rename = "@count", default)]
    count: u32,
    #[serde(rename = "numFmt", default)]
    num_fmt: Vec<NumFmt>,
}

#[derive(Debug, Deserialize, Serialize, Default)]
pub(crate) struct NumFmt {
    #[serde(rename = "@numFmtId", default)]
    num_fmt_id: u32,
    #[serde(rename = "@formatCode", default)]
    format_code: String
}