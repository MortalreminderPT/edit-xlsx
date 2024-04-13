use std::fmt::Formatter;
use serde::{Deserialize, Deserializer, Serialize, Serializer};
use serde::de::{Error, Visitor};

#[derive(Debug, Clone, PartialEq)]
pub(crate) enum RelType {
    Worksheets,
    Theme,
    Styles,
    Images,
    Hyperlinks,
    Drawings,
    VmlDrawing,
    Comments,
    Unknown,
    MetaData,
    SharedStrings,
    PrinterSettings,
    CalcChain,
    Table,
}

impl Serialize for RelType {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error> where S: Serializer {
        serializer.serialize_str(self.get_type())
    }
}

impl<'de> Visitor<'de> for RelType {
    type Value = RelType;

    fn expecting(&self, formatter: &mut Formatter) -> std::fmt::Result {
        todo!()
    }

    fn visit_str<E>(self, v: &str) -> Result<Self::Value, E> where E: Error {
        Ok(RelType::from_namespace(v))
    }
}
impl<'de> Deserialize<'de> for RelType {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error> where D: Deserializer<'de> {
        deserializer.deserialize_string(RelType::Unknown)
    }
}

impl RelType {
    pub(crate) fn get_type(&self) -> &str {
        match self {
            RelType::Worksheets => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            RelType::Theme => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
            RelType::Styles => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            RelType::Images => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            RelType::Hyperlinks => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            RelType::Drawings => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
            RelType::SharedStrings => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
            RelType::MetaData => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata",
            RelType::VmlDrawing => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
            RelType::Comments => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
            RelType::PrinterSettings => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings",
            RelType::CalcChain => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
            RelType::Table => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
            RelType::Unknown => "",
        }
    }

    fn from_namespace(namespace: &str) -> Self {
        match namespace {
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" => RelType::Worksheets,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" => RelType::Theme,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" => RelType::Styles,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" => RelType::Images,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" => RelType::Hyperlinks,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" => RelType::Drawings,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" => RelType::MetaData,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" => RelType::SharedStrings,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" => RelType::VmlDrawing,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" => RelType::Comments,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" => RelType::PrinterSettings,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" => RelType::CalcChain,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" => RelType::Table,
            &_ => RelType::Unknown
        }
    }
}