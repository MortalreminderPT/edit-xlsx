use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize, PartialEq, Clone)]
pub(crate) struct Alignment {
    #[serde(rename = "@horizontal", skip_serializing_if = "Option::is_none")]
    pub(crate) horizontal: Option<String>,
    #[serde(rename = "@vertical", skip_serializing_if = "Option::is_none")]
    pub(crate) vertical: Option<String>,
}

impl Alignment {
    pub(crate) fn default() -> Alignment {
        Alignment {
            horizontal: None,
            vertical: None,
        }
    }

    pub(crate) fn new(horizontal: Option<String>, vertical: Option<String>) -> Alignment {
        Alignment {
            horizontal,
            vertical,
        }
    }
}