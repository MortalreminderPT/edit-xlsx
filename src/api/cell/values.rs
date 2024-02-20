use std::fmt::Display;
use serde::{Deserialize, Deserializer, Serialize, Serializer};

#[derive(Debug, Serialize, Deserialize)]
#[serde(untagged)]
pub(crate) enum CellType {
    Boolean,
    Date,
    Error,
    Number,
    SharedString,
    String,
}

impl CellType {
    fn from_str(t: &str) -> CellType {
        match t {
            "b" => CellType::Boolean,
            "d" => CellType::Date,
            "e" => CellType::Error,
            "n" => CellType::Number,
            "s" => CellType::SharedString,
            _ => CellType::String,
        }
    }

    pub(crate) fn default() -> CellType {
        CellType::SharedString
    }

    fn as_str(&self) -> &str {
        match self {
            CellType::Boolean => "b",
            CellType::Date => "d",
            CellType::Error => "e",
            CellType::Number => "n",
            CellType::SharedString => "s",
            CellType::String => "str"
        }
    }

    pub(crate) fn se<S>(&self, serializer: S) -> Result<S::Ok, S::Error> where S: Serializer {
        serializer.serialize_str(self.as_str())
    }

    pub(crate) fn de<'de, D>(deserializer: D) -> Result<CellType, D::Error> where D: Deserializer<'de> {
        let s: &str = Deserialize::deserialize(deserializer).unwrap_or("s");
        Ok(CellType::from_str(s))
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