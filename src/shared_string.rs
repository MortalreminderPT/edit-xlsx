use std::io::{Read, Write};
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::PhoneticPr;

#[derive(Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename="sst")]
pub(crate) struct SharedString {
    #[serde(rename = "@count")]
    count: u32,
    #[serde(rename="@uniqueCount")]
    unique_count: u32,
    #[serde(rename="@xmlns")]
    xmlns: String,
    #[serde(rename="si")]
    pub string_item: Vec<StringItem>,
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub(crate) struct StringItem {
    #[serde(rename = "t")]
    pub text: String,
    #[serde(rename = "phoneticPr", skip_serializing_if = "Option::is_none")]
    phonetic_pr: Option<PhoneticPr>,
}

impl SharedString {
    pub(crate) fn add_text(&mut self, text: &str) -> usize {
        let item = StringItem { text: String::from(text), phonetic_pr: None };
        for i in 0..self.string_item.len() {
            if self.string_item[i] == item {
                return i;
            }
        }
        self.count += 1;
        self.unique_count += 1;
        self.string_item.push(item);
        self.string_item.len() - 1
    }
}

impl SharedString {
    pub fn from_path<P: AsRef<Path>>(file_path: P) -> SharedString {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::SharedStringFile).unwrap();
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let shared_string = de::from_str(&xml).unwrap();
        shared_string
    }

    pub fn save<P: AsRef<Path>>(&self, file_path: P) {
        let xml = se::to_string(&self).unwrap();
        // let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::SharedStringFile).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}