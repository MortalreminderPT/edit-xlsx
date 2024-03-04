use std::io;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::Properties;

#[derive(Debug, Deserialize, Serialize)]
#[serde(rename(serialize = "cp:coreProperties", deserialize = "coreProperties"))]
pub(crate) struct CoreProperties {
    #[serde(rename = "@xmlns:cp", skip_serializing_if = "Option::is_none")]
    xmlns_cp: Option<String>,
    #[serde(rename = "@xmlns:dc", skip_serializing_if = "Option::is_none")]
    xmlns_dc: Option<String>,
    #[serde(rename = "@xmlns:dcterms", skip_serializing_if = "Option::is_none")]
    xmlns_dcterms: Option<String>,
    #[serde(rename = "@xmlns:dcmitype", skip_serializing_if = "Option::is_none")]
    xmlns_dcmitype: Option<String>,
    #[serde(rename = "@xmlns:xsi", skip_serializing_if = "Option::is_none")]
    xmlns_xsi: Option<String>,
    #[serde(rename(serialize = "dc:title", deserialize = "title"), skip_serializing_if = "Option::is_none")]
    dc_title: Option<String>,
    #[serde(rename(serialize = "dc:subject", deserialize = "subject"), skip_serializing_if = "Option::is_none")]
    dc_subject: Option<String>,
    #[serde(rename(serialize = "dc:creator", deserialize = "creator"), skip_serializing_if = "Option::is_none")]
    dc_creator: Option<String>,
    #[serde(rename(serialize = "cp:keywords", deserialize = "keywords"), skip_serializing_if = "Option::is_none")]
    cp_keywords: Option<String>,
    #[serde(rename(serialize = "dc:description", deserialize = "description"), skip_serializing_if = "Option::is_none")]
    dc_description: Option<String>,
    #[serde(rename(serialize = "cp:lastModifiedBy", deserialize = "lastModifiedBy"), skip_serializing_if = "Option::is_none")]
    cp_last_modified_by: Option<String>,
    #[serde(rename(serialize = "dcterms:created", deserialize = "created"), skip_serializing_if = "Option::is_none")]
    dcterms_created: Option<PropertiesTime>,
    #[serde(rename(serialize = "dcterms:modified", deserialize = "modified"), skip_serializing_if = "Option::is_none")]
    dcterms_modified: Option<PropertiesTime>,
    #[serde(rename(serialize = "cp:category", deserialize = "category"), skip_serializing_if = "Option::is_none")]
    cp_category: Option<String>,
    #[serde(rename(serialize = "cp:contentStatus", deserialize = "contentStatus"), skip_serializing_if = "Option::is_none")]
    cp_content_status: Option<String>,
}

#[derive(Debug, Deserialize, Serialize)]
struct PropertiesTime {
    #[serde(rename(serialize = "@xsi:type", deserialize = "@type"), skip_serializing_if = "Option::is_none")]
    xsi_type: Option<String>,
    #[serde(rename = "$value", skip_serializing_if = "Option::is_none")]
    time: Option<String>,
}

impl CoreProperties {
    pub(crate) fn update_by_properties(&mut self, properties: &Properties) {
        if let Some(title) = properties.title {
            self.dc_title = Some(String::from(title));
        }
        if let Some(subject) = properties.subject {
            self.dc_subject = Some(String::from(subject));
        }
        if let Some(author) = properties.author {
            self.dc_creator = Some(String::from(author));
        }
        if let Some(category) = properties.category {
            self.cp_category = Some(String::from(category));
        }
        if let Some(keywords) = properties.keywords {
            self.cp_keywords = Some(String::from(keywords));
        }
        if let Some(comments) = properties.comments {
            self.dc_description = Some(String::from(comments));
        }
        if let Some(status) = properties.status {
            self.cp_content_status = Some(String::from(status));
        }
    }
}

impl CoreProperties {
    pub(crate) fn from_path<P: AsRef<Path>>(file_path: P) -> io::Result<CoreProperties> {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::CoreProperties)?;
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let properties: CoreProperties = de::from_str(&xml).unwrap();
        Ok(properties)
    }

    pub(crate) fn save<P: AsRef<Path>>(&self, file_path: P) {
        let xml = se::to_string_with_root("cp:coreProperties", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::CoreProperties).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}