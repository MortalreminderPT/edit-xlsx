use std::io;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::Properties;

#[derive(Debug, Deserialize, Serialize)]
#[serde(rename(serialize = "Properties", deserialize = "Properties"))]
pub(crate) struct AppProperties {
    #[serde(rename = "@xmlns", skip_serializing_if = "Option::is_none")]
    xmlns: Option<String>,
    #[serde(rename = "@xmlns:vt", skip_serializing_if = "Option::is_none")]
    xmlns_vt: Option<String>,
    #[serde(rename = "Application", skip_serializing_if = "Option::is_none")]
    application: Option<String>,
    #[serde(rename = "DocSecurity", skip_serializing_if = "Option::is_none")]
    doc_security: Option<u8>,
    #[serde(rename = "ScaleCrop", skip_serializing_if = "Option::is_none")]
    scale_crop: Option<bool>,
    #[serde(rename = "HeadingPairs", skip_serializing_if = "Option::is_none")]
    heading_pairs: Option<HeadingPairs>,
    #[serde(rename = "TitlesOfParts", skip_serializing_if = "Option::is_none")]
    titles_of_parts: Option<TitlesOfParts>,
    #[serde(rename = "Manager", skip_serializing_if = "Option::is_none")]
    manager: Option<String>,
    #[serde(rename = "Company", skip_serializing_if = "Option::is_none")]
    company: Option<String>,
    #[serde(rename = "LinksUpToDate", skip_serializing_if = "Option::is_none")]
    links_up_to_date: Option<bool>,
    #[serde(rename = "SharedDoc", skip_serializing_if = "Option::is_none")]
    shared_doc: Option<bool>,
    #[serde(rename = "HyperlinksChanged", skip_serializing_if = "Option::is_none")]
    hyperlinks_changed: Option<bool>,
    #[serde(rename = "AppVersion", skip_serializing_if = "Option::is_none")]
    app_version: Option<String>
}

impl AppProperties {
    pub(crate) fn update_by_properties(&mut self, properties: &Properties) {
        if let Some(manager) = properties.manager {
            self.manager = Some(String::from(manager));
        }
        if let Some(company) = properties.company {
            self.company = Some(String::from(company));
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct HeadingPairs {
    #[serde(rename(serialize = "vt:vector", deserialize = "vector"))]
    vt_vector: VtVector,
}

#[derive(Debug, Deserialize, Serialize)]
struct TitlesOfParts {
    #[serde(rename(serialize = "vt:vector", deserialize = "vector"))]
    vt_vector: VtVector,
}

#[derive(Debug, Deserialize, Serialize)]
struct VtVector {
    #[serde(rename = "@size")]
    size: u32,
    #[serde(rename = "@baseType", skip_serializing_if = "String::is_empty")]
    base_type: String,
    #[serde(rename(serialize = "vt:lpstr", deserialize = "lpstr"), default, skip_serializing_if = "Vec::is_empty")]
    vt_lpstr: Vec<VtLpstr>,
    #[serde(rename(serialize = "vt:variant", deserialize = "variant"), default, skip_serializing_if = "Vec::is_empty")]
    vt_variant: Vec<VtVariant>
}

#[derive(Debug, Deserialize, Serialize)]
struct VtVariant {
    #[serde(rename(serialize = "vt:lpstr", deserialize = "lpstr"), default, skip_serializing_if = "Vec::is_empty")]
    vt_lpstr: Vec<VtLpstr>,
    #[serde(rename(serialize = "vt:i4", deserialize = "i4"), default, skip_serializing_if = "Vec::is_empty")]
    vt_i4: Vec<VtI4>
}

#[derive(Debug, Deserialize, Serialize)]
struct VtLpstr {
    #[serde(rename = "$value")]
    value: String
}

#[derive(Debug, Deserialize, Serialize)]
struct VtI4 {
    #[serde(rename = "$value")]
    value: String
}

impl AppProperties {
    pub(crate) fn from_path<P: AsRef<Path>>(file_path: P) -> io::Result<AppProperties> {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::AppProperties)?;
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let properties: AppProperties = de::from_str(&xml).unwrap();
        Ok(properties)
    }

    pub(crate) fn save<P: AsRef<Path>>(&self, file_path: P) {
        let xml = se::to_string_with_root("Properties", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::AppProperties).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}