mod rel_type;
mod rel;

use std::collections::HashMap;
use serde::{Deserialize, Serialize};
use std::io;
use std::path::Path;
use quick_xml::{de, se};
use crate::api::relationship::Rel;
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::relationships::rel::RelationShip;
use crate::xml::relationships::rel_type::RelType;

#[derive(Debug, Clone, Deserialize, Serialize)]
pub(crate) struct Relationships {
    #[serde(rename = "@xmlns")]
    xmlns: String,
    #[serde(rename = "Relationship", default)]
    relationship: Vec<RelationShip>,
    #[serde(skip)]
    pub(crate) targets: Targets,
}

#[derive(Debug, Clone, Default)]
struct Targets {
    target: HashMap<String, Vec<u32>>,
}

impl Targets {
    fn add_target(&mut self, rel_type: &RelType, name: &str) {
        let name = Path::new(name).file_stem().unwrap().to_str().unwrap();
        let id = name.chars().filter(|&c| c >= '0' && c <= '9').collect::<String>().parse().unwrap_or(1);
        let mut target = self.target.get_mut(rel_type.get_type());
        let mut vec = Vec::new();
        let ids = target.get_or_insert(&mut vec);
        ids.push(id);
    }

    fn next_target(&mut self, rel_type: RelType) -> String {
        let key = rel_type.get_type();
        let value = self.target.get_mut(key);
        let max_id: u32 = match value {
            None => {
                let vec = vec![1];
                self.target.insert(key.to_string(), vec);
                1
            }
            Some(ids) => {
                let &max_id = ids.iter().max().unwrap_or(&1);
                ids.push(max_id + 1);
                max_id + 1
            }
        };
        let id = max_id + 1;// max_id + 1;
        match rel_type {
            RelType::Worksheets => format!("worksheets/sheet{id}.xml"),
            RelType::Theme => format!("theme/theme{id}.xml"),
            RelType::Styles => String::from("styles.xml"),
            RelType::Images => format!("../media/image{id}.png"),
            RelType::Hyperlinks => { "".to_string() }
            RelType::Drawings => format!("../drawings/drawing{id}.xml"),
            RelType::MetaData => "metadata.xml".to_string(),
            RelType::CalcChain => "calcChain.xml".to_string(),
            RelType::SharedStrings => { "".to_string() }
            RelType::PrinterSettings => { "".to_string() }
            RelType::VmlDrawing => { "".to_string() }
            RelType::Comments => { "".to_string() }
            RelType::Unknown => { "".to_string() }
        }
    }
}

unsafe impl Sync for Relationships {}

unsafe impl Send for Relationships {}

impl Default for Relationships {
    fn default() -> Self {
        Relationships {
            xmlns: "http://schemas.openxmlformats.org/package/2006/relationships".to_string(),
            relationship: vec![],
            targets: Default::default(),
        }
    }
}

impl Relationships {
    fn next_id(&self) -> u32 {
        1 + self.relationship.len() as u32
    }

    pub(crate) fn get_drawings_rid(&self) -> Option<u32> {
        let binding = self.get_target_by_type(RelType::Drawings);
        let targets = binding.first();
        match targets {
            Some(targets) => {
                let id: u32 = targets.chars().filter(|&c| c >= '0' && c <= '9').collect::<String>().parse().unwrap();
                Some(id)
            }
            None => {
                None
            }
        }
        // let a = self.relationship
        //     .iter()
        //     .filter(|r| r.rel_type == RelType::Drawings);
        // // rid
        // let target = self.get_target(rid);
    }
    
    pub(crate) fn get_vml_drawing_rid(&self) -> Option<u32> {
        self.get_rid_by_type(RelType::VmlDrawing).first().copied()
    }

    pub(crate) fn get_target(&self, r_id: &Rel) -> &String {
        self.relationship.iter()
            .find(|r| r.id == *r_id)
            .map(|r| &r.target)
            .unwrap()
    }

    // pub(crate) fn list_targets(&self, r_ids: Vec<Rel>) -> Vec<&String> {
    //     self.relationship.iter().filter(
    //         |r| r_ids.contains(&r.id)
    //     ).map(|r| &r.target).collect()
    // }

    fn get_rid_by_type(&self, rel_type: RelType) -> Vec<u32> {
        self.relationship
            .iter()
            .filter(|r| r.rel_type == rel_type)
            .map(|r| r.id.get_id())
            .collect()
    }

    fn get_target_by_type(&self, rel_type: RelType) -> Vec<String> {
        self.relationship
            .iter()
            .filter(|r| r.rel_type == rel_type)
            .map(|r| r.target.clone())
            .collect()
    }

    fn exist_type(&self, rel_type: RelType) -> bool {
        self.relationship
            .iter()
            .filter(|r| r.rel_type == rel_type)
            .count() > 0
    }

    pub(crate) fn add_worksheet(&mut self, id: u32) -> (u32, String) {
        let r_id = self.next_id();
        let target = self.targets.next_target(RelType::Worksheets);
        let rel = RelationShip::new(r_id, RelType::Worksheets, &target, None);
        self.relationship.push(rel);
        (r_id, target)
    }

    pub(crate) fn add_image(&mut self, id: u32, image_extension: &str) -> u32 {
        let r_id = self.next_id();
        self.relationship.push(RelationShip::new_image(r_id, id, image_extension));
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
    pub(crate) async fn from_path_async<P: AsRef<Path>>(file_path: P, rel_type: XlsxFileType) -> io::Result<Relationships> {
        Self::from_path(file_path, rel_type)
    }

    pub(crate) async fn save_async<P: AsRef<Path>>(&self, file_path: P, rel_type: XlsxFileType) {
        self.save(file_path, rel_type)
    }

    pub(crate) fn from_path<P: AsRef<Path>>(file_path: P, rel_type: XlsxFileType) -> io::Result<Relationships> {
        let mut file = XlsxFileReader::from_path(file_path, rel_type)?;
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let mut rel: Relationships = de::from_str(&xml).unwrap();
        rel.relationship.iter()
            .for_each(|r| rel.targets.add_target(&r.rel_type, &r.target));
        Ok(rel)
    }

    pub(crate) fn save<P: AsRef<Path>>(&self, file_path: P, rel_type: XlsxFileType) {
        let xml = se::to_string_with_root("Relationships", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, rel_type).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}