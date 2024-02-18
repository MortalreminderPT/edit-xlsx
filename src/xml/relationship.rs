mod rels;

use serde::{Deserialize, Serialize};
pub(crate) const SHEET_TYPE_STRING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
pub(crate) const THEME_TYPE_STRING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
pub(crate) const STYLES_TYPE_STRING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
pub(crate) const IMAGE_TYPE_STRING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
pub(crate) const HYPERLINK_TYPE_STRING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub(crate) struct  RelationShip {
    #[serde(rename = "@Id")]
    id: String,
    #[serde(rename = "@Type")]
    pub(crate) rel_type: String,
    #[serde(rename = "@Target")]
    target: String,
    #[serde(rename = "@TargetMode", skip_serializing_if = "Option::is_none")]
    target_mode: Option<String>
}

impl RelationShip {
    pub(crate) fn styles_default() -> RelationShip {
        RelationShip {
            id: "rId3".to_string(),
            rel_type: STYLES_TYPE_STRING.to_string(),
            target: "styles.xml".to_string(),
            target_mode: None,
        }
    }

    pub(crate) fn theme_default() -> RelationShip {
        RelationShip {
            id: "rId2".to_string(),
            rel_type: THEME_TYPE_STRING.to_string(),
            target: "theme/theme1.xml".to_string(),
            target_mode: None,
        }
    }

    pub(crate) fn sheet_default() -> RelationShip {
        RelationShip {
            id: "rId1".to_string(),
            rel_type: SHEET_TYPE_STRING.to_string(),
            target: "worksheets/sheet1.xml".to_string(),
            target_mode: None,
        }
    }

    pub(crate) fn new_sheet(r_id: u32, sheet_id: u32) -> RelationShip {
        RelationShip {
            id: format!("rId{}", r_id),
            rel_type: SHEET_TYPE_STRING.to_string(),
            target: format!("worksheets/sheet{sheet_id}.xml"),
            target_mode: None,
        }
    }

    pub(crate) fn new_theme(id: u32, offset: u32) -> RelationShip {
        RelationShip {
            id: format!("rId{}", id + offset),
            rel_type: THEME_TYPE_STRING.to_string(),
            target: format!("theme/theme{id}.xml"),
            target_mode: None,
        }
    }

    pub(crate) fn new_styles(id: u32, offset: u32) -> RelationShip {
        RelationShip {
            id: format!("rId{}", id + offset),
            rel_type: STYLES_TYPE_STRING.to_string(),
            target: String::from("styles.xml"),
            target_mode: None,
        }
    }

    pub(crate) fn new_image(r_id: u32, id: u32) -> RelationShip {
        RelationShip {
            id: format!("rId{}", r_id),
            rel_type: IMAGE_TYPE_STRING.to_string(),
            target: format!("../media/image{id}.png"),
            target_mode: None,
        }
    }

    pub(crate) fn new_hyperlink(r_id: u32, target: &str) -> RelationShip {
        RelationShip {
            id: format!("rId{}", r_id),
            rel_type: HYPERLINK_TYPE_STRING.to_string(),
            target: String::from(target),
            target_mode: Some(String::from("External")),
        }
    }
}
