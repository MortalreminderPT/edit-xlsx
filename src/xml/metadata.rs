use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::extension::{AddExtension, ExtensionList, ExtensionType};
use crate::xml::io::Io;

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Metadata {
    #[serde(rename = "@xmlns")]
    xmlns: String,
    #[serde(rename(serialize = "@xmlns:xda", deserialize = "@xmlns:xda"), default, skip_serializing_if = "String::is_empty")]
    xmlns_xda: String,
    #[serde(rename = "metadataTypes", skip_serializing_if = "Option::is_none")]
    metadata_types: Option<MetadataTypes>,
    #[serde(rename = "futureMetadata", skip_serializing_if = "Option::is_none")]
    future_metadata: Option<FutureMetadata>,
    #[serde(rename = "cellMetadata", skip_serializing_if = "Option::is_none")]
    cell_metadata: Option<CellMetadata>,
}

impl Default for Metadata {
    fn default() -> Self {
        Self {
            xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string(),
            xmlns_xda: "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray".to_string(),
            metadata_types: Default::default(),
            future_metadata: Default::default(),
            cell_metadata: Default::default(),
        }
    }
}

impl AddExtension for Metadata {
    fn add_extension(&mut self, e: ExtensionType) {
        let mut future_metadata = self.future_metadata.take();
        if let None = future_metadata {
            future_metadata = Some(FutureMetadata::default());
        }
        let mut future_metadata = future_metadata.unwrap();
        if future_metadata.bk.is_empty() {
            future_metadata.bk.push(Bk::default());
        }
        let extension_list = &mut future_metadata.bk[0].ext_lst;
        extension_list.add_extension(e);
        self.future_metadata = Some(future_metadata)
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct MetadataTypes {
    #[serde(rename = "@count", default)]
    count: u32,
    #[serde(rename = "metadataType", skip_serializing_if = "Vec::is_empty")]
    metadata_type: Vec<MetadataType>,
}
impl Default for MetadataTypes {
    fn default() -> Self {
        Self {
            count: 1,
            metadata_type: vec![MetadataType::default()],
        }
    }
}
#[derive(Debug, Deserialize, Serialize)]
struct MetadataType {
    #[serde(rename = "@name")]
    name: String,
    #[serde(rename = "@minSupportedVersion")]
    min_supported_version: u32,
    #[serde(rename = "@copy")]
    copy: u8,
    #[serde(rename = "@pasteAll")]
    paste_all: u8,
    #[serde(rename = "@pasteValues")]
    paste_values: u8,
    #[serde(rename = "@merge")]
    merge: u8,
    #[serde(rename = "@splitFirst")]
    split_first: u8,
    #[serde(rename = "@rowColShift")]
    row_col_shift: u8,
    #[serde(rename = "@clearFormats")]
    clear_formats: u8,
    #[serde(rename = "@clearComments")]
    clear_comments: u8,
    #[serde(rename = "@assign")]
    assign: u8,
    #[serde(rename = "@coerce")]
    coerce: u8,
    #[serde(rename = "@cellMeta", skip_serializing_if = "Option::is_none")]
    cell_meta: Option<u8>,
}
impl Default for MetadataType {
    fn default() -> Self {
        Self {
            name: "XLDAPR".to_string(),
            min_supported_version: 120000,
            copy: 1,
            paste_all: 1,
            paste_values: 1,
            merge: 1,
            split_first: 1,
            row_col_shift: 1,
            clear_formats: 1,
            clear_comments: 1,
            assign: 1,
            coerce: 1,
            cell_meta: Some(1),
        }
    }
}
#[derive(Debug, Deserialize, Serialize)]
struct FutureMetadata {
    #[serde(rename = "@name", default)]
    name: String,
    #[serde(rename = "@count", default)]
    count: u32,
    #[serde(rename = "bk", skip_serializing_if = "Vec::is_empty")]
    bk: Vec<Bk>
}
impl Default for FutureMetadata {
    fn default() -> Self {
        Self {
            name: "XLDAPR".to_string(),
            count: 1,
            bk: vec![Default::default()],
        }
    }
}
#[derive(Debug, Deserialize, Serialize)]
struct CellMetadata {
    #[serde(rename = "@count", default)]
    count: u32,
    #[serde(rename = "bk", skip_serializing_if = "Vec::is_empty")]
    bk: Vec<Bk>
}

impl Default for CellMetadata {
    fn default() -> Self {
        Self {
            count: 1,
            bk: vec![Bk::default_rc()],
        }
    }
}

#[derive(Debug, Deserialize, Serialize, Default)]
struct Bk {
    #[serde(rename = "extLst", default, skip_serializing_if = "ExtensionList::is_empty")]
    ext_lst: ExtensionList,
    #[serde(rename = "rc", default, skip_serializing_if = "Vec::is_empty")]
    rc: Vec<Rc>
}

impl Bk {
    fn default_rc() -> Bk {
        Bk {
            ext_lst: Default::default(),
            rc: vec![Default::default()],
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct Rc {
    #[serde(rename = "@t")]
    t: u8,
    #[serde(rename = "@v")]
    v: u8,
}
impl Default for Rc {
    fn default() -> Self {
        Self {
            t: 1,
            v: 0,
        }
    }
}

impl Io<Metadata> for Metadata {
    fn from_path<P: AsRef<Path>>(file_path: P) -> std::io::Result<Metadata> {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::MetaData)?;
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let metadata = de::from_str(&xml).unwrap();
        Ok(metadata)
    }

    fn save<P: AsRef<Path>>(& self, file_path: P) {
        let xml = se::to_string_with_root("metadata", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::MetaData).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}