mod rel_type;
mod rel;

use std::collections::HashMap;
use std::fs::File;
use serde::{Deserialize, Serialize};
use std::io::Read;
use std::path::Path;
use quick_xml::{de, se};
use zip::read::ZipFile;
use zip::ZipArchive;
use crate::api::relationship::Rel;
use crate::file::{XlsxFileType, XlsxFileWriter};
use crate::xml::relationships::rel::RelationShip;
use crate::xml::relationships::rel_type::RelType;
use crate::xml::workbook::Workbook;

#[derive(Debug, Clone, Deserialize, Serialize)]
pub(crate) struct Relationships {
    #[serde(rename = "@xmlns")]
    xmlns: String,
    #[serde(rename = "Relationship", default)]
    relationship: Vec<RelationShip>,
    #[serde(skip)]
    targets: Targets,
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

    fn next_target(&mut self, rel_type: RelType, id: u32) -> String {
        // let key = rel_type.get_type();
        // let value = self.target.get_mut(key);
        // let max_id: u32 = match value {
        //     None => {
        //         let vec = vec![1];
        //         self.target.insert(key.to_string(), vec);
        //         1
        //     }
        //     Some(ids) => {
        //         let &max_id = ids.iter().max().unwrap_or(&1);
        //         ids.push(max_id + 1);
        //         max_id + 1
        //     }
        // };
        // let id = max_id + 1;// max_id + 1;
        match rel_type {
            RelType::Worksheets => format!("worksheets/sheet{id}.xml"),
            RelType::Theme => format!("theme/theme{id}.xml"),
            RelType::Styles => String::from("styles.xml"),
            RelType::Images => format!("../media/image{id}.png"),
            RelType::Drawings => format!("../drawings/drawing{id}.xml"),
            RelType::Hyperlinks => { "".to_string() }
            RelType::MetaData => "metadata.xml".to_string(),
            RelType::CalcChain => "calcChain.xml".to_string(),
            RelType::Table => format!("../tables/table{id}.xml"),
            RelType::Chart => format!("../charts/chart{id}.xml"),
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
    pub(crate) fn next_id(&self) -> u32 {
        1 + self.relationship.len() as u32
    }

    ///
    /// Add a worksheet in workbook.rel and return the target id and rid
    ///
    pub(crate) fn add_worksheet_v2(&mut self) -> (u32, u32) {
        let r_id = self.next_id();
        let sheet_target_id: u32 = 1 + self.relationship
            .iter()
            .filter(|r| r.rel_type == RelType::Worksheets)
            .map(|r| r.target[16..r.target.len() - 4].parse().unwrap_or(0))
            .max()
            .unwrap_or(0);
        let rel = RelationShip::new(r_id, RelType::Worksheets, &format!("worksheets/sheet{sheet_target_id}.xml"), None);
        self.relationship.push(rel);
        (r_id, sheet_target_id)
    }

    pub(crate) fn next_sheet_target_id(&self) -> u32 {
        1 + self.relationship.iter()
            .filter(|r| r.rel_type == RelType::Worksheets)
            .map(|r| r.target.chars().filter(|&c| c >= '0' && c <= '9').collect::<String>().parse().unwrap_or(0))
            .max()
            .unwrap_or(0)
    }

    pub(crate) fn get_drawings_rids(&self) -> Vec<u32> {
        let binding = self.get_target_by_type(RelType::Drawings);
        let rids: Vec<u32> = binding.iter()
            .map(|s|s.chars().filter(|&c| c >= '0' && c <= '9').collect::<String>().parse().unwrap())
            .collect();
        rids
    }
    
    pub(crate) fn get_vml_drawing_rid(&self) -> Option<u32> {
        self.get_rid_by_type(RelType::VmlDrawing).first().copied()
    }

    pub(crate) fn get_target(&self, r_id: &Rel) -> (&String, u32) {
        let target = self.relationship.iter()
            .find(|r| r.id == *r_id)
            .map(|r| &r.target)
            .unwrap();
        let target_id: u32 = target[16..target.len() - 4].parse().unwrap();
        (target, target_id)
    }

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
        let target = self.targets.next_target(RelType::Worksheets, id);
        let rel = RelationShip::new(r_id, RelType::Worksheets, &target, None);
        self.relationship.push(rel);
        (r_id, target.clone())
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
    // pub(crate) async fn from_path_async<P: AsRef<Path>>(file_path: P, rel_type: XlsxFileType) -> io::Result<Relationships> {
    //     Self::from_path(file_path, rel_type)
    // }

    pub(crate) async fn save_async<P: AsRef<Path>>(&self, file_path: P, rel_type: XlsxFileType) {
        self.save(file_path, rel_type)
    }

    // pub(crate) fn from_path<P: AsRef<Path>>(file_path: P, rel_type: XlsxFileType) -> io::Result<Relationships> {
    //     let mut file = XlsxFileReader::from_path(file_path, rel_type)?;
    //     let mut xml = String::new();
    //     file.read_to_string(&mut xml).unwrap();
    //     let mut rel: Relationships = de::from_str(&xml).unwrap();
    //     rel.relationship.iter()
    //         .for_each(|r| rel.targets.add_target(&r.rel_type, &r.target));
    //     Ok(rel)
    // }

    pub(crate) fn save<P: AsRef<Path>>(&self, file_path: P, rel_type: XlsxFileType) {
        let xml = se::to_string_with_root("Relationships", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, rel_type).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}