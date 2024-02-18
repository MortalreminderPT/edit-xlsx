use std::io;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::io::Io;
use crate::xml::relationship::{RelationShip};

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Relationships {
    #[serde(rename = "@xmlns")]
    xmlns: String,
    #[serde(rename = "Relationship")]
    relationship: Vec<RelationShip>,
}

impl Default for Relationships {
    fn default() -> Self {
        Relationships {
            xmlns: "http://schemas.openxmlformats.org/package/2006/relationships".to_string(),
            relationship: vec![RelationShip::styles_default(), RelationShip::theme_default(), RelationShip::sheet_default()],
        }
    }
}

impl Relationships {
    pub(crate) fn next_id(&self) -> u32 {
        1 + self.relationship.len() as u32
    }

    pub(crate) fn add_worksheet(&mut self, r_id: u32, id: u32) {
        self.relationship.push(RelationShip::new_sheet(r_id, id));
    }

    fn add_hyperlink(&mut self, r_id: u32, target: &str) {
        self.relationship.push(RelationShip::new_hyperlink(r_id, target));
    }
}

impl Io<Relationships> for Relationships {
    fn from_path<P: AsRef<Path>>(file_path: P) -> io::Result<Relationships> {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::WorkbookRels)?;
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let rel: Relationships = de::from_str(&xml).unwrap();
        Ok(rel)
    }

    fn save<P: AsRef<Path>>(&mut self, file_path: P) {
        let xml = se::to_string_with_root("Relationships", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::WorkbookRels).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}