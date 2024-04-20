use std::fs::File;
use std::io;
use std::io::Read;
use std::path::Path;
use quick_xml::de;
use serde::Deserialize;
use zip::ZipArchive;
use crate::xml::content_types::ContentTypes;
use crate::xml::drawings::Drawings;
use crate::xml::drawings::vml_drawing::VmlDrawing;
use crate::xml::metadata::Metadata;
use crate::xml::relationships::Relationships;
use crate::xml::shared_string::SharedString;
use crate::xml::style::StyleSheet;
use crate::xml::workbook::Workbook;
use crate::xml::worksheet::WorkSheet;

pub(crate) trait Io<T: Default> {
    fn save<P: AsRef<Path>>(&self, file_path: P);
    async fn save_async<P: AsRef<Path>>(&self, file_path: P) {
        self.save(file_path)
    }
}

pub(crate) trait IoV2<T: for<'de> Deserialize<'de> + Default> {
    fn from_zip_file(archive: &mut ZipArchive<File>, path: &str) -> Option<T> {
        if let Ok(mut file) = archive.by_name(path) {
            let mut xml = String::new();
            file.read_to_string(&mut xml).unwrap();
            let result = de::from_str(&xml).unwrap_or_default();
            Some(result)
        } else {
            None
        }
    }
}

impl IoV2<Workbook> for Workbook {}
impl IoV2<WorkSheet> for WorkSheet{}
impl IoV2<StyleSheet> for StyleSheet {}
impl IoV2<ContentTypes> for ContentTypes{}
impl IoV2<Relationships> for Relationships{}
impl IoV2<Metadata> for Metadata{}
impl IoV2<SharedString> for SharedString{}
impl IoV2<Drawings> for Drawings{}
impl IoV2<VmlDrawing> for VmlDrawing{}