mod xda_dynamic_array_properties;
mod x15_workbook_pr;
mod x14_slicer_styles;
mod x15_timeline_styles;

use std::collections::HashSet;
use std::hash::{Hash, Hasher};
use serde::{Deserialize, Serialize};
use crate::xml::common::XmlnsAttrs;
use crate::xml::extension::x14_slicer_styles::X14SlicerStyles;
use crate::xml::extension::x15_timeline_styles::X15TimelineStyles;
use crate::xml::extension::x15_workbook_pr::X15WorkbookPr;
use crate::xml::extension::xda_dynamic_array_properties::XdaDynamicArrayProperties;

pub(crate) enum ExtensionType {
    X15WorkbookPr,
    X14SlicerStyles,
    X15TimelineStyles,
    XdaDynamicArrayProperties,
}

pub(crate) trait AddExtension {
    fn add_extension(&mut self, e: ExtensionType);
}

#[derive(Debug, Deserialize, Serialize, Default)]
pub(crate) struct ExtensionList {
    #[serde(rename = "ext")]
    ext: HashSet<Extension>,
}

impl ExtensionList {
    pub(crate) fn is_empty(&self) -> bool {
        self.ext.is_empty()
    }
}

impl AddExtension for ExtensionList {
    fn add_extension(&mut self, e: ExtensionType) {
        self.ext.insert(Extension::from_extension_type(e));
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct Extension {
    #[serde(rename = "@uri")]
    uri: String,
    #[serde(rename = "@xmlns:x14", skip_serializing_if = "Option::is_none")]
    xmlns_x14: Option<String>,
    #[serde(rename = "@xmlns:x15", skip_serializing_if = "Option::is_none")]
    xmlns_x15: Option<String>,
    #[serde(rename(serialize = "x15:workbookPr", deserialize = "workbookPr"), skip_serializing_if = "Option::is_none")]
    x15_workbook_pr: Option<X15WorkbookPr>,
    #[serde(rename(serialize = "x14:slicerStyles", deserialize = "slicerStyles"), skip_serializing_if = "Option::is_none")]
    x14_slicer_styles: Option<X14SlicerStyles>,
    #[serde(rename(serialize = "x15:timelineStyles", deserialize = "timelineStyles"), skip_serializing_if = "Option::is_none")]
    x15_timeline_styles: Option<X15TimelineStyles>,
    #[serde(rename(serialize = "xda:dynamicArrayProperties", deserialize = "dynamicArrayProperties"), skip_serializing_if = "Option::is_none")]
    xda_dynamic_array_properties: Option<XdaDynamicArrayProperties>,
}

impl PartialEq for Extension {
    fn eq(&self, other: &Self) -> bool {
        self.uri.eq(&other.uri)
    }
}

impl Eq for Extension {}

impl Hash for Extension {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.uri.hash(state)
    }
}

impl Extension {
    fn from_extension_type(extension_type: ExtensionType) -> Self {
        match extension_type {
            ExtensionType::X15WorkbookPr => Self::new_x15_workbook_pr(),
            ExtensionType::X14SlicerStyles => Self::new_x14_slicer_styles(),
            ExtensionType::X15TimelineStyles => Self::new_x15_timeline_styles(),
            ExtensionType::XdaDynamicArrayProperties => Self::new_xda_dynamic_array_properties(),
        }
    }

    fn new_x15_workbook_pr() -> Self {
        Self {
            uri: "".to_string(),
            xmlns_x14: None,
            xmlns_x15: Some("http://schemas.microsoft.com/office/spreadsheetml/2010/11/main".to_string()),
            x15_workbook_pr: Some(Default::default()),
            x14_slicer_styles: None,
            x15_timeline_styles: None,
            xda_dynamic_array_properties: None,
        }
    }

    fn new_x14_slicer_styles() -> Self {
        Self {
            uri: "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}".to_string(),
            xmlns_x14: Some("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main".to_string()),
            xmlns_x15: None,
            x15_workbook_pr: None,
            x14_slicer_styles: Some(Default::default()),
            x15_timeline_styles: None,
            xda_dynamic_array_properties: None,
        }
    }

    fn new_x15_timeline_styles() -> Self {
        Self {
            uri: "{9260A510-F301-46a8-8635-F512D64BE5F5}".to_string(),
            xmlns_x14: None,
            xmlns_x15: Some("http://schemas.microsoft.com/office/spreadsheetml/2010/11/main".to_string()),
            x15_workbook_pr: None,
            x14_slicer_styles: None,
            x15_timeline_styles: Some(Default::default()),
            xda_dynamic_array_properties: None,
        }
    }

    fn new_xda_dynamic_array_properties() -> Self {
        Self {
            uri: "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}".to_string(),
            xmlns_x14: None,
            xmlns_x15: None,
            x15_workbook_pr: None,
            x14_slicer_styles: None,
            x15_timeline_styles: None,
            xda_dynamic_array_properties: Some(Default::default()),
        }
    }
}