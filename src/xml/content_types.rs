use std::collections::HashSet;
use std::fs::File;
use std::hash::Hash;
use std::io;
use std::io::Read;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use zip::read::ZipFile;
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::io::Io;
use crate::xml::relationships::Relationships;

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

impl ContentType {
    fn get_extension(&self) -> Option<&str> {
        if let ContentType::Default { extension, content_type } = self {
            Some(&extension)
        } else {
            None
        }
    }
}

impl ContentTypes {
    fn get_mut_by_extension(&self, extension: &str) -> bool {
        self.content_types.iter().find(|c| c.get_extension() == Some(extension)).is_some()
    }
}

impl ContentTypes {
    pub(crate) fn add_png(&mut self) {
        self.content_types.insert(ContentType::png_default());
    }
    pub(crate) fn add_bin(&mut self, extension: &str) {
        if self.get_mut_by_extension(extension) {
            return;
        }
        self.content_types.insert(ContentType::octet_stream_default(extension));
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

    fn octet_stream_default(extension: &str) -> ContentType {
        ContentType::Default {
            extension: extension.to_string(),
            content_type: "application/octet-stream".to_string(),
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

impl ContentTypes {
    pub(crate) fn from_file(file: &File) -> ContentTypes {
        let mut xml = String::new();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        let file_path = "[Content_Types].xml";
        let content_types = match archive.by_name(&file_path) {
            Ok(mut file) => {
                file.read_to_string(&mut xml).unwrap();
                de::from_str(&xml).unwrap()
            }
            Err(_) => {
                ContentTypes::default()
            }
        };
        content_types
    }
}

impl Io<ContentTypes> for ContentTypes {
    fn from_zip_file(mut file: &mut ZipFile) -> Self {
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        de::from_str(&xml).unwrap_or_default()
    }

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
