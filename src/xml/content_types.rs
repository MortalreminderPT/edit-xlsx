use std::collections::HashSet;
use std::hash::Hash;
use std::io;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::io::Io;

#[derive(Debug, Deserialize, Serialize, Default)]
pub(crate) struct ContentTypes {
    #[serde(rename = "@xmlns")]
    xmlns: String,
    #[serde(rename = "$value")]
    content_types: HashSet<ContentType>,
}

#[derive(Debug, Deserialize, Serialize, PartialEq, Eq, Hash)]
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

impl ContentTypes {
    pub(crate) fn add_png(&mut self) {
        self.content_types.insert(ContentType::png_default());
    }
    pub(crate) fn add_drawing(&mut self, id: u32) { self.content_types.insert(ContentType::drawing_override(id)); }
    pub(crate) fn add_metadata(&mut self) { self.content_types.insert(ContentType::metadata_override()); }
}

impl ContentType {
    fn png_default() -> ContentType {
        ContentType::Default {
            extension: "png".to_string(),
            content_type: "image/png".to_string(),
        }
    }

    fn drawing_override(id: u32) -> ContentType {
        ContentType::Override {
            part_name: format!("/xl/drawings/drawing{id}.xml"),
            content_type: "application/vnd.openxmlformats-officedocument.drawing+xml".to_string(),
        }
    }

    fn metadata_override() -> ContentType {
        ContentType::Override {
            part_name: "/xl/metadata.xml".to_string(),
            content_type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml".to_string(),
        }
    }
}

impl Io<ContentTypes> for ContentTypes {
    fn from_path<P: AsRef<Path>>(file_path: P) -> io::Result<ContentTypes> {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::ContentTypes)?;
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let types: ContentTypes = de::from_str(&xml).unwrap();
        Ok(types)
    }

    fn save<P: AsRef<Path>>(& self, file_path: P) {
        let xml = se::to_string_with_root("Types", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::ContentTypes).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}
