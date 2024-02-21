extern crate proc_macro;
use std::hash::Hash;
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

impl XmlnsAttrs {
    fn default_none() -> XmlnsAttrs {
        XmlnsAttrs {
            xmlns: None,
            xmlns_r: None,
            xmlns_mc: None,
            mc_ignorable: None,
            xmlns_x14: None,
            xmlns_x14ac: None,
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
}
// 
// impl Extension {
//     fn default_dynamic_array() -> Self {
//         Self {
//             uri: "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}".to_string(),
//             xmlns_attrs: XmlnsAttrs::default_none(),
//             x15_workbook_pr: None,
//             x14_slicer_styles: None,
//             x15_timeline_styles: None,
//             dynamic_array_properties: Some(XdaDynamicArrayProperties::default()),
//         }
//     }
// }

#[derive(Debug, Deserialize, Serialize, Clone, PartialEq, Eq, Hash)]
pub(crate) struct Element<T: Clone + PartialEq + Eq + Hash> {
    #[serde(rename = "@val")]
    val: T
}

impl<T: Clone + PartialEq + Eq + Hash> Element<T> {
    pub(crate) fn from_val(val: T) -> Element<T> {
        Element {
            val
        }
    }
}

impl<T:Clone + PartialEq + Eq + Hash + Default> Default for Element<T> {
    fn default() -> Self {
        Element {
            val: T::default(),
        }
    }
}

impl<T: Clone + PartialEq + Eq + Hash + Default> FromFormat<T> for Element<T> {
    fn set_attrs_by_format(&mut self, format: &T) {
        self.val = format.clone();
    }
}

impl XmlnsAttrs {
    pub(crate) fn workbook_default() -> XmlnsAttrs {
        XmlnsAttrs {
            xmlns: Some("http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string()),
            xmlns_r: Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships".to_string()),
            xmlns_mc: Some("http://schemas.openxmlformats.org/markup-compatibility/2006".to_string()),
            mc_ignorable: Some("x15".to_string()),
            xmlns_x14: None,
            xmlns_x14ac: None,
            xmlns_x15: Some("http://schemas.microsoft.com/office/spreadsheetml/2010/11/main".to_string()),
            xmlns_xr: None,
            xmlns_xr6: None,
            xmlns_xr10: None,
            xmlns_xr2: None,
            xmlns_xr3: None,
            xmlns_x16r2: None,
            xr_uid: None,
        }
    }
    
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
    
    pub(crate) fn shared_string_default() -> XmlnsAttrs {
        XmlnsAttrs {
            xmlns: Some("http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string()),
            xmlns_r: None,
            xmlns_mc: None,
            mc_ignorable: None,
            xmlns_x14: None,
            xmlns_x14ac: None,
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
}

pub(crate) fn is_zero(num: &u32) -> bool {
    num.eq(&0)
}
pub(crate) trait FromFormat<T>: Default {
    fn set_attrs_by_format(&mut self, format: &T);
    fn from_format(format: &T) -> Self {
        let mut def = Self::default();
        def.set_attrs_by_format(format);
        def
    }
}