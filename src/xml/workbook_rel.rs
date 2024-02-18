use std::io;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::io::Io;
const SHEET_TYPE_STRING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const THEME_TYPE_STRING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const STYLES_TYPE_STRING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Relationships {
    #[serde(rename = "@xmlns")]
    xmlns: String,
    #[serde(rename = "Relationship")]
    relationship: Vec<RelationShip>,
    #[serde(skip)]
    last_sheet_id: u32,
    #[serde(skip)]
    sheet_offset: u32,
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
    fn styles_default() -> RelationShip {
        RelationShip {
            id: "rId3".to_string(),
            rel_type: STYLES_TYPE_STRING.to_string(),
            target: "styles.xml".to_string(),
        }
    }

    fn theme_default() -> RelationShip {
        RelationShip {
            id: "rId2".to_string(),
            rel_type: THEME_TYPE_STRING.to_string(),
            target: "theme/theme1.xml".to_string(),
        }
    }

    fn sheet_default() -> RelationShip {
        RelationShip {
            id: "rId1".to_string(),
            rel_type: SHEET_TYPE_STRING.to_string(),
            target: "worksheets/sheet1.xml".to_string(),
        }
    }

    fn new_sheet(id: u32, offset: u32) -> RelationShip {
        RelationShip {
            id: format!("rId{}", id + offset),
            rel_type: SHEET_TYPE_STRING.to_string(),
            target: format!("worksheets/sheet{id}.xml"),
        }
    }

    fn new_theme(id: u32, offset: u32) -> RelationShip {
        RelationShip {
            id: format!("rId{}", id + offset),
            rel_type: THEME_TYPE_STRING.to_string(),
            target: format!("theme/theme{id}.xml"),
        }
    }

    fn new_styles(id: u32, offset: u32) -> RelationShip {
        RelationShip {
            id: format!("rId{}", id + offset),
            rel_type: STYLES_TYPE_STRING.to_string(),
            target: String::from("styles.xml"),
        }
    }
}

impl Default for Relationships {
    fn default() -> Self {
        Relationships {
            xmlns: "http://schemas.openxmlformats.org/package/2006/relationships".to_string(),
            relationship: vec![RelationShip::styles_default(), RelationShip::theme_default(), RelationShip::sheet_default()],
            last_sheet_id: 0,
            sheet_offset: 0,
        }
    }
}

impl Relationships {
    pub(crate) fn update(&mut self, worksheet_size: u32, theme_size: u32, style_sheet_size: u32) {
        self.relationship = Vec::new();
        let mut offset = 0;
        for id in 1..=worksheet_size {
            self.relationship.push(
                RelationShip::new_sheet(id, offset)
            )
        }
        offset += worksheet_size;
        for id in worksheet_size..=theme_size {
            self.relationship.push(
                RelationShip::new_theme(id, offset)
            )
        }
        offset += theme_size;
        for id in theme_size..=style_sheet_size {
            self.relationship.push(
                RelationShip::new_styles(id, offset)
            )
        }
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