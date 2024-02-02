use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct PhoneticPr {
    #[serde(rename = "@fontId")]
    font_id: u32,
    #[serde(rename = "@type")]
    phonetic_pr_type: String
}

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct XmlnsAttrs {
    #[serde(rename = "@xmlns", skip_serializing_if = "Option::is_none")]
    xmlns: Option<String>,
    #[serde(rename = "@xmlns:r", skip_serializing_if = "Option::is_none")]
    xmlns_r: Option<String>,
    #[serde(rename = "@xmlns:mc", skip_serializing_if = "Option::is_none")]
    xmlns_mc: Option<String>,
    #[serde(rename(serialize = "@mc:Ignorable", deserialize = "@Ignorable"), skip_serializing_if = "Option::is_none")]
    mc_ignorable: Option<String>,
    #[serde(rename = "@xmlns:x14ac", skip_serializing_if = "Option::is_none")]
    xmlns_x14ac: Option<String>,
    #[serde(rename = "@xmlns:x15", skip_serializing_if = "Option::is_none")]
    xmlns_x15: Option<String>,
    #[serde(rename = "@xmlns:xr", skip_serializing_if = "Option::is_none")]
    xmlns_xr: Option<String>,
    #[serde(rename = "@xmlns:xr6", skip_serializing_if = "Option::is_none")]
    xmlns_xr6: Option<String>,
    #[serde(rename = "@xmlns:xr10", skip_serializing_if = "Option::is_none")]
    xmlns_xr10: Option<String>,
    #[serde(rename = "@xmlns:xr2", skip_serializing_if = "Option::is_none")]
    xmlns_xr2: Option<String>,
    #[serde(rename = "@xmlns:xr3", skip_serializing_if = "Option::is_none")]
    xmlns_xr3: Option<String>,
    #[serde(rename(serialize = "@xr:uid", deserialize = "@uid"), skip_serializing_if = "Option::is_none")]
    xr_uid: Option<String>,
}

impl XmlnsAttrs {
    pub(crate) fn default() -> XmlnsAttrs {
        XmlnsAttrs {
            xmlns: Some("http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string()),
            xmlns_r: Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships".to_string()),
            xmlns_mc: Some("http://schemas.openxmlformats.org/markup-compatibility/2006".to_string()),
            mc_ignorable: Some("x14ac".to_string()),
            xmlns_x14ac: Some("http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac".to_string()),
            xmlns_x15: None,
            xmlns_xr: None,
            xmlns_xr6: None,
            xmlns_xr10: None,
            xmlns_xr2: None,
            xmlns_xr3: None,
            xr_uid: None,
        }
    }
}