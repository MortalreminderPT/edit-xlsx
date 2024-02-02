use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::manage::XmlIo;

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Relationships {
    #[serde(rename = "@xmlns")]
    xmlns: String,
    #[serde(rename = "Relationship")]
    relationship: Vec<RelationShip>
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
    fn new_sheet(sheet_id: u32) -> RelationShip {
        RelationShip {
            id: format!("rId{sheet_id}"),
            rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet".to_string(),
            target: format!("worksheets/sheet{sheet_id}.xml"),
        }
    }
}

impl Relationships {
    pub(crate) fn add_worksheet(&mut self) -> u32 {
        let id = match self.last_sheet_id() {
            None => 1,
            Some(max_id) => 1 + max_id
        };
        self.relationship.push(
            RelationShip::new_sheet(id)
        );
        self.relationship.iter_mut()
            .filter(|rel| { !rel.target.starts_with("worksheets") })
            .for_each(|rel| {
                let new_id = 1 + &rel.id[3..].parse::<u32>().unwrap();
                rel.id = format!("rId{new_id}");
            });
        id
    }
    pub(crate) fn last_sheet_id(&self) -> Option<u32> {
        self.relationship
            .iter()
            .filter(|rel| { rel.target.starts_with("worksheets") })
            .map(|r| r.id[3..].parse::<u32>().unwrap())
            .max()
    }
}

impl XmlIo<Relationships> for Relationships {
    fn from_path<P: AsRef<Path>>(file_path: P) -> Relationships {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::WorkbookRels).unwrap();
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let rel = de::from_str(&xml).unwrap();
        rel
    }

    fn save<P: AsRef<Path>>(&mut self, file_path: P) {
        let xml = se::to_string_with_root("Relationships", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::WorkbookRels).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}