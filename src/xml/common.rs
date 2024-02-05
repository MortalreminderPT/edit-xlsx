extern crate proc_macro;
use proc_macro::TokenStream;
use serde::{Deserialize, Deserializer, Serialize};

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
    #[serde(rename = "@xmlns:x14", skip_serializing_if = "Option::is_none")]
    xmlns_x14: Option<String>,
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
    #[serde(rename = "@xmlns:x16r2", skip_serializing_if = "Option::is_none")]
    xmlns_x16r2: Option<String>,
    #[serde(rename(serialize = "@xr:uid", deserialize = "@uid"), skip_serializing_if = "Option::is_none")]
    xr_uid: Option<String>,
}

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct ExtLst {
    #[serde(rename = "ext")]
    ext: Vec<Ext>,
}

#[derive(Debug, Deserialize, Serialize)]
struct Ext {
    #[serde(rename = "@uri")]
    uri: String,
    #[serde(flatten)]
    xmlns_attrs: XmlnsAttrs,
    #[serde(rename(serialize = "x15:workbookPr", deserialize = "workbookPr"), skip_serializing_if = "Option::is_none")]
    x15_workbook_pr: Option<X15WorkbookPr>,
    #[serde(rename(serialize = "x14:slicerStyles", deserialize = "slicerStyles"), skip_serializing_if = "Option::is_none")]
    x14_slicer_styles: Option<X14SlicerStyles>,
    #[serde(rename(serialize = "x15:timelineStyles", deserialize = "timelineStyles"), skip_serializing_if = "Option::is_none")]
    x15_timeline_styles: Option<X15TimelineStyles>,
}

#[derive(Debug, Deserialize, Serialize)]
struct X15WorkbookPr {
    #[serde(rename = "@chartTrackingRefBase", skip_serializing_if = "Option::is_none")]
    chart_tracking_ref_base: Option<u32>
}

#[derive(Debug, Deserialize, Serialize)]
struct X14SlicerStyles {
    #[serde(rename = "@defaultSlicerStyle", skip_serializing_if = "Option::is_none")]
    default_slicer_style: Option<String>
}

#[derive(Debug, Deserialize, Serialize)]
struct X15TimelineStyles {
    #[serde(rename = "@defaultTimelineStyle", skip_serializing_if = "Option::is_none")]
    default_timeline_style: Option<String>
}

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Element<T> {
    #[serde(rename = "@val")]
    val: T
}

impl<T> Element<T> {
    pub(crate) fn from_val(val: T) -> Element<T> {
        Element {
            val
        }
    }
}

impl XmlnsAttrs {
    pub(crate) fn worksheet_default() -> XmlnsAttrs {
        XmlnsAttrs {
            xmlns: Some("http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string()),
            xmlns_r: Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships".to_string()),
            xmlns_mc: Some("http://schemas.openxmlformats.org/markup-compatibility/2006".to_string()),
            mc_ignorable: Some("x14ac".to_string()),
            xmlns_x14: None,
            xmlns_x14ac: Some("http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac".to_string()),
            xmlns_x15: None,
            xmlns_xr: None,
            xmlns_xr6: None,
            xmlns_xr10: None,
            xmlns_xr2: None,
            xmlns_xr3: None,
            xmlns_x16r2: None,
            xr_uid: None,
        }
    }

    pub(crate) fn stylesheet_default() -> XmlnsAttrs {
        XmlnsAttrs {
            xmlns: Some("http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string()),
            xmlns_r: Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships".to_string()),
            xmlns_mc: Some("http://schemas.openxmlformats.org/markup-compatibility/2006".to_string()),
            mc_ignorable: Some("x14ac x16r2".to_string()),
            xmlns_x14: None,
            xmlns_x14ac: Some("http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac".to_string()),
            xmlns_x15: None,
            xmlns_xr: None,
            xmlns_xr6: None,
            xmlns_xr10: None,
            xmlns_xr2: None,
            xmlns_xr3: None,
            xmlns_x16r2: Some("http://schemas.microsoft.com/office/spreadsheetml/2015/02/main".to_string()),
            xr_uid: None,
        }
    }
}