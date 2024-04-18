use std::io;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileType, XlsxFileWriter};
use crate::xml::namespaces::office as o;
use crate::xml::namespaces::vml as v;

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct VmlDrawing {
    #[serde(rename(serialize = "@xmlns:v", deserialize = "@xmlns:v"))]
    xmlns_v: String,
    #[serde(rename(serialize = "@xmlns:o", deserialize = "@xmlns:o"))]
    xmlns_o: String,
    #[serde(rename(serialize = "@xmlns:x", deserialize = "@xmlns:x"))]
    xmlns_x: String,
    #[serde(rename(serialize = "o:shapelayout", deserialize = "shapelayout"))]
    shapelayout: o::ShapeLayout,
    #[serde(rename(serialize = "v:shapetype", deserialize = "shapetype"))]
    shapetype: v::ShapeType,
    #[serde(rename(serialize = "v:shape", deserialize = "shape"))]
    shape: v::Shape,
}

impl VmlDrawing {
    // pub(crate) fn from_path<P: AsRef<Path>>(file_path: P, drawing_id: u32) -> io::Result<VmlDrawing> {
    //     let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::VmlDrawings(drawing_id))?;
    //     let mut xml = String::new();
    //     file.read_to_string(&mut xml).unwrap();
    //     let drawing: VmlDrawing = de::from_str(&xml).unwrap();
    //     Ok(drawing)
    // }

    pub(crate) fn save<P: AsRef<Path>>(&self, file_path: P, drawing_id: u32) {
        let xml = se::to_string_with_root("xml", &self).unwrap();
        // let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::VmlDrawings(drawing_id)).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}
