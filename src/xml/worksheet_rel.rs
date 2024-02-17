use std::io;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::manage::XmlIo;
const IMAGE_TYPE_STRING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Relationships {
    #[serde(rename = "@xmlns")]
    xmlns: String,
    #[serde(rename = "Relationship")]
    relationship: Vec<RelationShip>,
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
    fn new_image(id: u32, offset: u32) -> RelationShip {
        RelationShip {
            id: format!("rId{}", id + offset),
            rel_type: IMAGE_TYPE_STRING.to_string(),
            target: format!("../media/image{id}.png"),
        }
    }
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
    pub(crate) fn add_image(&mut self, id: u32) {
        self.relationship.push(
            RelationShip::new_image(id, 0)
        )
    }

    pub(crate) fn update(&mut self, image_size: u32) {
        self.relationship = Vec::new();
        let mut offset = 0;
        for id in 1..=image_size {
            self.relationship.push(
                RelationShip::new_image(id, offset)
            )
        }
    }
}

impl Relationships {
    pub(crate) fn from_path<P: AsRef<Path>>(file_path: P, sheet_id: u32) -> io::Result<Relationships> {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::WorksheetRels(sheet_id))?;
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let rel: Relationships = de::from_str(&xml).unwrap();
        Ok(rel)
    }

    pub(crate) fn save<P: AsRef<Path>>(&mut self, file_path: P, sheet_id: u32) {
        let xml = se::to_string_with_root("Relationships", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::WorksheetRels(sheet_id)).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}