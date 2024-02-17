use std::collections::HashSet;
use std::hash::{Hash, Hasher};
use std::io;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::manage::XmlIo;

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct ContentTypes {
    #[serde(rename = "@xmlns")]
    xmlns: String,
    #[serde(rename = "$value")]
    content_types: HashSet<ContentType>,
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
enum ContentType{
    Default {
        #[serde(rename = "@Extension")]
        extension: String,
        #[serde(rename = "@ContentType")]
        content_type: String,
    },
    Override {
        #[serde(rename = "@PartName")]
        part_name: String,
        #[serde(rename = "@ContentType")]
        content_type: String,
    }
}

impl Eq for ContentType {}

impl Hash for ContentType {
    fn hash<H: Hasher>(&self, state: &mut H) {
        match self {
            ContentType::Default { extension, content_type } => {
                extension.hash(state)
            }
            ContentType::Override { part_name, content_type } => {
                part_name.hash(state)
            }
        }
    }
}

impl ContentTypes {
    pub(crate) fn add_png(&mut self) {
        self.content_types.insert(ContentType::png_default());
    }
}

impl ContentType {
    pub(crate) fn png_default() -> ContentType {
        ContentType::Default {
            extension: "png".to_string(),
            content_type: "image/png".to_string(),
        }
    }
}

impl XmlIo<ContentTypes> for ContentTypes {
    fn from_path<P: AsRef<Path>>(file_path: P) -> io::Result<ContentTypes> {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::ContentTypes)?;
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let types: ContentTypes = de::from_str(&xml).unwrap();
        Ok(types)
    }

    fn save<P: AsRef<Path>>(&mut self, file_path: P) {
        let xml = se::to_string_with_root("Types", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::ContentTypes).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}
