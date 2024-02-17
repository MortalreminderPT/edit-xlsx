use std::io;
use std::path::Path;
use quick_xml::de;
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType};
use crate::xml::manage::XmlIo;

const IMAGE_TYPE_STRING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Relationships {
    #[serde(rename = "@xmlns")]
    xmlns: String,
    #[serde(rename = "Relationship")]
    relationship: Vec<RelationShip>,
    #[serde(skip)]
    last_r_id: u32,
    #[serde(skip)]
    r_id_offset: u32,
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
struct RelationShip {
    #[serde(rename = "@Id")]
    id: String,
    #[serde(rename = "@Type")]
    rel_type: String,
    #[serde(rename = "@Target")]
    target: String,
}

impl RelationShip {
    fn new_image(id: u32, name: &str) -> RelationShip {
        RelationShip {
            id: format!("rId{id}"),
            rel_type: String::from(IMAGE_TYPE_STRING),
            target: format!("../media/{name}"),
        }
    }
}

impl Relationships {
    fn add_image(&mut self, name: &str) -> u32 {
        let id = self.last_r_id + 1;
        self.relationship.push(
            RelationShip::new_image(id, name)
        );
        self.last_r_id += 1;
        self.r_id_offset += 1;
        id
    }
}


impl Default for Relationships {
    fn default() -> Self {
        Relationships {
            xmlns: "http://schemas.openxmlformats.org/package/2006/relationships".to_string(),
            relationship: vec![],
            last_r_id: 0,
            r_id_offset: 0,
        }
    }
}
