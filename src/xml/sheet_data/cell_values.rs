use std::fmt::Display;
use serde::{Deserialize, Deserializer, Serialize, Serializer};

#[derive(Debug, Serialize, Deserialize)]
#[serde(untagged)]
pub(crate) enum CellValues {
    Boolean,
    Date,
    Error,
    Number,
    SharedString,
    String,
}

impl CellValues {
    fn from_str(t: &str) -> CellValues {
        match t {
            "b" => CellValues::Boolean,
            "d" => CellValues::Date,
            "e" => CellValues::Error,
            "n" => CellValues::Number,
            "s" => CellValues::SharedString,
            _ => CellValues::String,
        }
    }

    pub(crate) fn default() -> CellValues {
        CellValues::SharedString
    }

    fn as_str(&self) -> &str {
        match self {
            CellValues::Boolean => "b",
            CellValues::Date => "d",
            CellValues::Error => "e",
            CellValues::Number => "n",
            CellValues::SharedString => "s",
            CellValues::String => "str"
        }
    }

    pub(crate) fn se<S>(&self, serializer: S) -> Result<S::Ok, S::Error> where S: Serializer {
        serializer.serialize_str(self.as_str())
    }

    pub(crate) fn de<'de, D>(deserializer: D) -> Result<CellValues, D::Error> where D: Deserializer<'de> {
        let s: &str = Deserialize::deserialize(deserializer).unwrap_or("s");
        Ok(CellValues::from_str(s))
    }
}


pub(crate) trait CellDisplay {
    fn to_display(&self) -> String;
}

pub(crate) trait CellType {
    fn to_cell_val(&self) -> CellValues {
        CellValues::String
    }
}

impl<T: Display> CellDisplay for T {
    fn to_display(&self) -> String {
        self.to_string()
    }
}

impl CellType for &str {}
impl CellType for String {}

impl CellType for i8 {
    fn to_cell_val(&self) -> CellValues {
        CellValues::Number
    }
}
impl CellType for i16 {
    fn to_cell_val(&self) -> CellValues {
        CellValues::Number
    }
}
impl CellType for i32 {
    fn to_cell_val(&self) -> CellValues {
        CellValues::Number
    }
}
impl CellType for i64 {
    fn to_cell_val(&self) -> CellValues {
        CellValues::Number
    }
}
impl CellType for i128 {
    fn to_cell_val(&self) -> CellValues {
        CellValues::Number
    }
}
impl CellType for u8 {
    fn to_cell_val(&self) -> CellValues {
        CellValues::Number
    }
}
impl CellType for u16 {
    fn to_cell_val(&self) -> CellValues {
        CellValues::Number
    }
}
impl CellType for u32 {
    fn to_cell_val(&self) -> CellValues {
        CellValues::Number
    }
}
impl CellType for u64 {
    fn to_cell_val(&self) -> CellValues {
        CellValues::Number
    }
}
impl CellType for u128 {
    fn to_cell_val(&self) -> CellValues {
        CellValues::Number
    }
}
impl CellType for f32 {
    fn to_cell_val(&self) -> CellValues {
        CellValues::Number
    }
}
impl CellType for f64 {
    fn to_cell_val(&self) -> CellValues {
        CellValues::Number
    }
}

impl CellType for bool {
    fn to_cell_val(&self) -> CellValues {
        CellValues::Boolean
    }
}