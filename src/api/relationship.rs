use std::fmt::Formatter;
use serde::{Deserialize, Deserializer, Serialize, Serializer};
use serde::de::{Error, Visitor};

#[derive(Debug, Clone, Default, PartialEq)]
pub struct Rel {
    id: u32
}

impl Rel {
    pub(crate) fn from_str(r_id: &str) -> Rel {
        let id: u32 = r_id[3..].parse().unwrap();
        Rel {
            id,
        }
    }

    pub(crate) fn from_id(id: u32) -> Rel {
        Rel {
            id,
        }
    }

    pub(crate) fn get_id(&self) -> u32 {
        self.id
    }
}

impl<'de> Visitor<'de> for Rel {
    type Value = Rel;
    fn expecting(&self, formatter: &mut Formatter) -> std::fmt::Result {
        todo!()
    }
    fn visit_str<E>(self, v: &str) -> Result<Self::Value, E> where E: Error {
        let rel = Rel::from_str(v);
        Ok(rel)
    }
}

impl<'de> Deserialize<'de> for Rel {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error> where D: Deserializer<'de> {
        let rel = Rel::default();
        deserializer.deserialize_string(rel)
    }
}

impl Serialize for Rel {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error> where S: Serializer {
        serializer.serialize_str(&format!("rId{}", self.id))
    }
}