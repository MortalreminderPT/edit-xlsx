use serde::{Deserialize, Serialize};
use crate::api::relationship::Rel;
use crate::xml::relationships::rel_type::RelType;

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub(crate) struct RelationShip {
    #[serde(rename = "@Id")]
    pub(crate) id: Rel,
    #[serde(rename = "@Type")]
    pub(crate) rel_type: RelType,
    #[serde(rename = "@Target")]
    pub(crate) target: String,
    #[serde(rename = "@TargetMode", skip_serializing_if = "Option::is_none")]
    target_mode: Option<String>
}

impl RelationShip {
    pub(crate) fn new_sheet(r_id: u32, sheet_id: u32) -> RelationShip {
        RelationShip {
            id: Rel::from_id(r_id),
            rel_type: RelType::Worksheets,
            target: format!("worksheets/sheet{sheet_id}.xml"),
            target_mode: None,
        }
    }

    pub(crate) fn new_theme(id: u32) -> RelationShip {
        RelationShip {
            id: Rel::from_id(id),
            rel_type: RelType::Theme,
            target: format!("theme/theme{id}.xml"),
            target_mode: None,
        }
    }

    pub(crate) fn new_styles(id: u32) -> RelationShip {
        RelationShip {
            id: Rel::from_id(id),
            rel_type: RelType::Styles,
            target: String::from("styles.xml"),
            target_mode: None,
        }
    }

    pub(crate) fn new_image(r_id: u32, id: u32) -> RelationShip {
        RelationShip {
            id: Rel::from_id(r_id),
            rel_type: RelType::Images,
            target: format!("../media/image{id}.png"),
            target_mode: None,
        }
    }

    pub(crate) fn new_hyperlink(r_id: u32, target: &str) -> RelationShip {
        RelationShip {
            id: Rel::from_id(r_id),
            rel_type: RelType::Hyperlinks,
            target: String::from(target),
            target_mode: Some(String::from("External")),
        }
    }

    pub(crate) fn new_drawing(r_id: u32, id: u32) -> RelationShip {
        RelationShip {
            id: Rel::from_id(r_id),
            rel_type: RelType::Drawings,
            target: format!("../drawings/drawing{id}.xml"),
            target_mode: None,
        }
    }

    pub(crate) fn new_metadata(r_id: u32) -> RelationShip {
        RelationShip {
            id: Rel::from_id(r_id),
            rel_type: RelType::MetaData,
            target: "metadata.xml".to_string(),
            target_mode: None,
        }
    }
}