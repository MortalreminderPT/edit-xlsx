use std::io;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::manage::XmlIo;
const SHEET_TYPE_STRING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

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
            rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles".to_string(),
            target: "styles.xml".to_string(),
        }
    }

    fn theme_default() -> RelationShip {
        RelationShip {
            id: "rId2".to_string(),
            rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme".to_string(),
            target: "theme/theme1.xml".to_string(),
        }
    }

    fn sheet_default() -> RelationShip {
        RelationShip {
            id: "rId1".to_string(),
            rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet".to_string(),
            target: "worksheets/sheet1.xml".to_string(),
        }
    }

    fn new_sheet(sheet_id: u32) -> RelationShip {
        RelationShip {
            id: format!("rId{sheet_id}"),
            rel_type: String::from(SHEET_TYPE_STRING),
            target: format!("worksheets/sheet{sheet_id}.xml"),
        }
    }
}

impl Relationships {
    pub(crate) fn add_worksheet(&mut self) -> u32 {
        let id = self.last_sheet_id + 1;
        self.relationship.push(
            RelationShip::new_sheet(id)
        );
        self.last_sheet_id += 1;
        self.sheet_offset += 1;
        id
    }

    fn last_sheet_id(&self) -> Option<u32> {
        self.relationship
            .iter()
            .filter(|rel| { rel.target.starts_with("worksheets") })
            .map(|r| r.id[3..].parse::<u32>().unwrap())
            .max()
    }

    fn reset_offset(&mut self) {
        self.relationship
            .iter_mut()
            .filter(|rel| { !rel.target.starts_with("worksheets") })
            .for_each(|rel| {
                let new_id = self.sheet_offset + &rel.id[3..].parse::<u32>().unwrap();
                rel.id = format!("rId{new_id}");
            });
        self.sheet_offset = 0;
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

impl XmlIo<Relationships> for Relationships {
    fn from_path<P: AsRef<Path>>(file_path: P) -> io::Result<Relationships> {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::WorkbookRels)?;
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let mut rel: Relationships = de::from_str(&xml).unwrap();
        rel.last_sheet_id = rel.last_sheet_id().unwrap_or(0);
        rel.sheet_offset = 0;
        Ok(rel)
    }

    fn save<P: AsRef<Path>>(&mut self, file_path: P) {
        self.reset_offset();
        let xml = se::to_string_with_root("Relationships", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::WorkbookRels).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}