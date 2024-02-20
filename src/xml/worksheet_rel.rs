use std::io;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::relationship::RelationShip;

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

    pub(crate) fn add_image(&mut self, r_id: u32, id: u32) {
        self.relationship.push(
            RelationShip::new_image(r_id, id)
        )
    }

    pub(crate) fn add_hyperlink(&mut self, r_id: u32, target: &str) {
        self.relationship.push(RelationShip::new_hyperlink(r_id, target));
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