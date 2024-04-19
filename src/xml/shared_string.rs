use std::fs::File;
use std::io;
use std::io::Read;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::xml::worksheet::sheet_data::cell::inline_string::RichText;
use crate::xml::io::Io;

#[derive(Debug, Deserialize, Serialize, Clone)]
#[serde(rename="sst")]
pub(crate) struct SharedString {
    // #[serde(flatten)]
    // xmlns_attrs: XmlnsAttrs,
    // #[serde(rename = "@count", default)]
    // count: u32,
    // #[serde(rename = "@uniqueCount", default)]
    // unique_count: u32,
    #[serde(rename = "si", default = "Vec::new")]
    string_item: Vec<StringItem>,
}

impl Default for SharedString {
    fn default() -> Self {
        Self {
            // xmlns_attrs: XmlnsAttrs::shared_string_default(),
            // count: 0,
            // unique_count: 0,
            string_item: vec![],
        }
    }
}

#[derive(Debug, Deserialize, Serialize, Clone)]
pub(crate) struct StringItem {
    #[serde(rename = "t", default, skip_serializing_if = "String::is_empty")]
    text: String,
    #[serde(rename = "r", skip_serializing_if = "Vec::is_empty")]
    rich_texts: Vec<RichText>
    // #[serde(rename = "phoneticPr", skip_serializing_if = "Option::is_none")]
    // phonetic_pr: Option<PhoneticPr>,
}

impl SharedString {
    pub(crate) fn get_text(&self, id: usize) -> Option<&str> {
        match self.string_item.get(id) {
            Some(string_item) => Some(string_item.text.as_str()),
            None => None
        }
    }

    pub(crate) fn get_rich_text(&self, id: usize) -> Vec<RichText> {
        match self.string_item.get(id) {
            Some(string_item) => string_item.rich_texts.clone(),
            None => vec![]
        }
    }
    // pub(crate) fn add_text(&mut self, text: &str) -> u32 {
    //     let item = StringItem { text: String::from(text), phonetic_pr: None };
    //     for i in 0..self.string_item.len() {
    //         if self.string_item[i].text == item.text {
    //             return i as u32;
    //         }
    //     }
    //     self.count += 1;
    //     self.unique_count += 1;
    //     self.string_item.push(item);
    //     self.string_item.len() as u32 - 1
    // }
}

impl SharedString {

    pub(crate) fn from_file(file: &File) -> SharedString {
        let mut xml = String::new();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        let file_path = "xl/sharedStrings.xml";
        let shared_string = match archive.by_name(&file_path) {
            Ok(mut file) => {
                file.read_to_string(&mut xml).unwrap();
                de::from_str(&xml).unwrap()
            }
            Err(_) => {
                SharedString::default()
            }
        };
        shared_string
    }
}

impl Io<SharedString> for SharedString {
    fn save<P: AsRef<Path>>(&self, file_path: P) {
        return;
        // let xml = se::to_string_with_root("sst", &self).unwrap();
        // let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        // let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::SharedStringFile).unwrap();
        // file.write_all(xml.as_ref()).unwrap();
    }
}