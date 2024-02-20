mod rels;

use serde::{Deserialize, Serialize};
use crate::api::relationship::Rel;
use std::io;
use std::path::Path;
use quick_xml::{de, se};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};

pub(crate) const SHEET_TYPE_STRING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
pub(crate) const THEME_TYPE_STRING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
pub(crate) const STYLES_TYPE_STRING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
pub(crate) const IMAGE_TYPE_STRING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
pub(crate) const HYPERLINK_TYPE_STRING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
pub(crate) const DRAWING_TYPE_STRING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Relationships {
    #[serde(rename = "@xmlns")]
    xmlns: String,
    #[serde(rename = "Relationship", default)]
    relationship: Vec<RelationShip>,
}

impl Default for Relationships {
    fn default() -> Self {
        Relationships {
            xmlns: "http://schemas.openxmlformats.org/package/2006/relationships".to_string(),
            relationship: vec![],
        }
    }
}

impl Relationships {
    pub(crate) fn next_id(&self) -> u32 {
        1 + self.relationship.len() as u32
    }

    pub(crate) fn list_drawings(&self) -> Vec<u32> {
        let targets: Vec<u32> = self.relationship.iter()
            .filter(|r| { r.rel_type == DRAWING_TYPE_STRING })
            .map(|r| {
                r.target
                    .chars().filter(
                    |&c| {
                        c >= '0' && c <= '9'
                    }
                ).collect::<String>().parse::<u32>().unwrap()
            })
            .collect();
        targets
    }

    pub(crate) fn add_worksheet(&mut self, r_id: u32, id: u32) {
        self.relationship.push(RelationShip::new_sheet(r_id, id));
    }

    pub(crate) fn add_image(&mut self, r_id: u32, id: u32) {
        self.relationship.push(
            RelationShip::new_image(r_id, id)
        )
    }

    pub(crate) fn add_hyperlink(&mut self, r_id: u32, target: &str) {
        self.relationship.push(RelationShip::new_hyperlink(r_id, target));
    }


    pub(crate) fn add_drawing(&mut self, r_id: u32, id: u32) {
        self.relationship.push(
            RelationShip::new_drawing(r_id, id)
        )
    }
}

impl Relationships {
    pub(crate) fn from_path<P: AsRef<Path>>(file_path: P, rel_type: XlsxFileType) -> io::Result<Relationships> {
        let mut file = XlsxFileReader::from_path(file_path, rel_type)?;
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let rel: Relationships = de::from_str(&xml).unwrap();
        Ok(rel)
    }

    pub(crate) fn save<P: AsRef<Path>>(&mut self, file_path: P, rel_type: XlsxFileType) {
        let xml = se::to_string_with_root("Relationships", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, rel_type).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}


#[derive(Debug, Deserialize, Serialize, PartialEq)]
struct RelationShip {
    #[serde(rename = "@Id")]
    id: Rel,
    #[serde(rename = "@Type")]
    pub(crate) rel_type: String,
    #[serde(rename = "@Target")]
    pub(crate) target: String,
    #[serde(rename = "@TargetMode", skip_serializing_if = "Option::is_none")]
    target_mode: Option<String>
}

impl RelationShip {
    pub(crate) fn styles_default() -> RelationShip {
        RelationShip {
            id: Rel::from_id(3),
            rel_type: STYLES_TYPE_STRING.to_string(),
            target: "styles.xml".to_string(),
            target_mode: None,
        }
    }

    pub(crate) fn theme_default() -> RelationShip {
        RelationShip {
            id: Rel::from_id(2),
            rel_type: THEME_TYPE_STRING.to_string(),
            target: "theme/theme1.xml".to_string(),
            target_mode: None,
        }
    }

    pub(crate) fn sheet_default() -> RelationShip {
        RelationShip {
            id: Rel::from_id(1),
            rel_type: SHEET_TYPE_STRING.to_string(),
            target: "worksheets/sheet1.xml".to_string(),
            target_mode: None,
        }
    }

    pub(crate) fn new_sheet(r_id: u32, sheet_id: u32) -> RelationShip {
        RelationShip {
            id: Rel::from_id(r_id),
            rel_type: SHEET_TYPE_STRING.to_string(),
            target: format!("worksheets/sheet{sheet_id}.xml"),
            target_mode: None,
        }
    }

    pub(crate) fn new_theme(id: u32, offset: u32) -> RelationShip {
        RelationShip {
            id: Rel::from_id(id),
            rel_type: THEME_TYPE_STRING.to_string(),
            target: format!("theme/theme{id}.xml"),
            target_mode: None,
        }
    }

    pub(crate) fn new_styles(id: u32, offset: u32) -> RelationShip {
        RelationShip {
            id: Rel::from_id(id),
            rel_type: STYLES_TYPE_STRING.to_string(),
            target: String::from("styles.xml"),
            target_mode: None,
        }
    }

    pub(crate) fn new_image(r_id: u32, id: u32) -> RelationShip {
        RelationShip {
            id: Rel::from_id(r_id),
            rel_type: IMAGE_TYPE_STRING.to_string(),
            target: format!("../media/image{id}.png"),
            target_mode: None,
        }
    }

    pub(crate) fn new_hyperlink(r_id: u32, target: &str) -> RelationShip {
        RelationShip {
            id: Rel::from_id(r_id),
            rel_type: HYPERLINK_TYPE_STRING.to_string(),
            target: String::from(target),
            target_mode: Some(String::from("External")),
        }
    }

    pub(crate) fn new_drawing(r_id: u32, id: u32) -> RelationShip {
        RelationShip {
            id: Rel::from_id(r_id),
            rel_type: DRAWING_TYPE_STRING.to_string(),
            target: format!("../drawings/drawing{id}.xml"),
            target_mode: None,
        }
    }
}
