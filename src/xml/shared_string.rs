use std::io;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::common::{PhoneticPr, XmlnsAttrs};
use crate::xml::io::Io;

#[derive(Debug, Deserialize, Serialize)]
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

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct StringItem {
    #[serde(rename = "t", default)]
    text: String,
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

impl Io<SharedString> for SharedString {
    fn from_path<P: AsRef<Path>>(file_path: P) -> io::Result<SharedString> {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::SharedStringFile)?;
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let shared_string = de::from_str(&xml).unwrap();
        Ok(shared_string)
    }

    fn save<P: AsRef<Path>>(&self, file_path: P) {
        return;
        // let xml = se::to_string_with_root("sst", &self).unwrap();
        // let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        // let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::SharedStringFile).unwrap();
        // file.write_all(xml.as_ref()).unwrap();
    }
}