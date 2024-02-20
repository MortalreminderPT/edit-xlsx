use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct PageMargins {
    #[serde(rename = "@left")]
    left: f64,
    #[serde(rename = "@right")]
    right: f64,
    #[serde(rename = "@top")]
    top: f64,
    #[serde(rename = "@bottom")]
    bottom: f64,
    #[serde(rename = "@header")]
    header: f64,
    #[serde(rename = "@footer")]
    footer: f64,
}

impl Default for PageMargins {
    fn default() -> PageMargins {
        PageMargins {
            left: 0.7,
            right: 0.7,
            top: 0.75,
            bottom: 0.75,
            header: 0.3,
            footer: 0.3,
        }
    }
}