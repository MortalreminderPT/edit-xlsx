use std::fmt::{Display, Formatter};
use serde::{Deserialize, Deserializer, Serialize, Serializer};
use serde::de::{Error, Visitor};

#[derive(Debug, Clone, PartialEq)]
pub(crate) enum CellType {
    Boolean,
    Date,
    Error,
    Number,
    SharedString,
    String,
    Undefined,
    InlineString,
}

impl Default for CellType {
    fn default() -> Self {
        CellType::Undefined
    }
}

impl CellType {
    fn from_str(t: &str) -> CellType {
        match t {
            "b" => CellType::Boolean,
            "d" => CellType::Date,
            "e" => CellType::Error,
            "n" => CellType::Number,
            "s" => CellType::SharedString,
            "str" => CellType::String,
            "inlineStr" => CellType::InlineString,
            _ => CellType::Undefined,
        }
    }

    fn to_str(&self) -> &str {
        match self {
            CellType::Boolean => "b",
            CellType::Date => "d",
            CellType::Error => "e",
            CellType::Number => "n",
            CellType::SharedString => "s",
            CellType::InlineString => "inlineStr",
            CellType::String => "str",
            CellType::Undefined => "",
        }
    }

    pub(crate) fn de<'de, D>(deserializer: D) -> Result<CellType, D::Error> where D: Deserializer<'de> {
        let s: &str = Deserialize::deserialize(deserializer).unwrap_or("str");
        Ok(CellType::from_str(s))
    }
}

impl Serialize for CellType {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error> where S: Serializer {
        let str = self.to_str();
        serializer.serialize_str(str)
    }
}

impl<'de> Visitor<'de> for CellType {
    type Value = CellType;

    fn expecting(&self, formatter: &mut Formatter) -> std::fmt::Result {
        todo!()
    }

    fn visit_str<E>(self, v: &str) -> Result<Self::Value, E> where E: Error {
        Ok(CellType::from_str(v))
    }
}

impl<'de> Deserialize<'de> for CellType {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error> where D: Deserializer<'de> {
        let cell_type = CellType::Undefined;
        deserializer.deserialize_str(cell_type)
    }
}


/// 可以被展示的Cell文字内容
pub(crate) trait CellDisplay {
    fn to_display(&self) -> String;
}

/// Cell文字内容的类型
pub(crate) trait CellValue {
    fn to_cell_type(&self) -> CellType {
        CellType::String
    }
}

impl<T: Display> CellDisplay for T {
    fn to_display(&self) -> String {
        self.to_string()
    }
}

impl CellValue for &str {}
impl CellValue for String {}
impl CellValue for i8 {
    fn to_cell_type(&self) -> CellType {
        CellType::Number
    }
}
impl CellValue for i16 {
    fn to_cell_type(&self) -> CellType {
        CellType::Number
    }
}
impl CellValue for i32 {
    fn to_cell_type(&self) -> CellType {
        CellType::Number
    }
}
impl CellValue for i64 {
    fn to_cell_type(&self) -> CellType {
        CellType::Number
    }
}
impl CellValue for i128 {
    fn to_cell_type(&self) -> CellType {
        CellType::Number
    }
}
impl CellValue for u8 {
    fn to_cell_type(&self) -> CellType {
        CellType::Number
    }
}
impl CellValue for u16 {
    fn to_cell_type(&self) -> CellType {
        CellType::Number
    }
}
impl CellValue for u32 {
    fn to_cell_type(&self) -> CellType {
        CellType::Number
    }
}
impl CellValue for u64 {
    fn to_cell_type(&self) -> CellType {
        CellType::Number
    }
}
impl CellValue for u128 {
    fn to_cell_type(&self) -> CellType {
        CellType::Number
    }
}
impl CellValue for f32 {
    fn to_cell_type(&self) -> CellType {
        CellType::Number
    }
}
impl CellValue for f64 {
    fn to_cell_type(&self) -> CellType {
        CellType::Number
    }
}
impl CellValue for bool {
    fn to_cell_type(&self) -> CellType {
        CellType::Boolean
    }
}