mod rel_type;
mod rel;

use serde::{Deserialize, Serialize};
use std::io;
use std::path::Path;
use quick_xml::{de, se};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::relationships::rel::RelationShip;
use crate::xml::relationships::rel_type::RelType;

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
    fn next_id(&self) -> u32 {
        1 + self.relationship.len() as u32
    }

    pub(crate) fn get_drawings_rid(&self) -> Option<u32> {
        self.get_rid_by_type(RelType::Drawings).first().copied()
    }
    
    pub(crate) fn get_vml_drawing_rid(&self) -> Option<u32> {
        self.get_rid_by_type(RelType::VmlDrawing).first().copied()
    }

    fn get_rid_by_type(&self, rel_type: RelType) -> Vec<u32> {
        self.relationship
            .iter()
            .filter(|r| r.rel_type == rel_type)
            .map(|r| r.id.get_id())
            .collect()
    }

    fn exist_type(&self, rel_type: RelType) -> bool {
        self.relationship
            .iter()
            .filter(|r| r.rel_type == rel_type)
            .count() > 0
    }

    pub(crate) fn add_worksheet(&mut self, id: u32) -> u32 {
        let r_id = self.next_id();
        self.relationship.push(RelationShip::new_sheet(r_id, id));
        r_id
    }

    pub(crate) fn add_image(&mut self, id: u32) -> u32 {
        let r_id = self.next_id();
        self.relationship.push(RelationShip::new_image(r_id, id));
        r_id
    }

    pub(crate) fn add_hyperlink(&mut self, target: &str) -> u32 {
        let r_id = self.next_id();
        self.relationship.push(RelationShip::new_hyperlink(r_id, target));
        r_id
    }

    pub(crate) fn add_drawings(&mut self, id: u32) -> u32 {
        let r_id = self.next_id();
        self.relationship.push(RelationShip::new_drawing(r_id, id));
        r_id
    }

    pub(crate) fn get_or_add_metadata(&mut self) -> u32 {
        let r_id = self.get_rid_by_type(RelType::MetaData);
        if r_id.is_empty() {
            let r_id = self.next_id();
            self.relationship.push(RelationShip::new_metadata(r_id));
            return r_id;
        }
        return r_id[0]
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
