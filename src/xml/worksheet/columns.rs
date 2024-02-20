use serde::{Deserialize, Serialize};
use crate::result::ColResult;

#[derive(Debug, Deserialize, Serialize, PartialEq, Default)]
pub(crate) struct Cols {
    col: Vec<Col>
}

impl Cols {
    pub(crate) fn add_col(&mut self, min: u32, max: u32, width: Option<f64>, style: Option<u32>, best_fit: Option<u8>) -> ColResult<()> {
        let mut col = Col::new(min, max, 1, width, style, best_fit);
        if let None = width {
            col.custom_width = 0;
        }
        self.col.push(col);
        Ok(())
    }
    pub(crate) fn is_empty(&self) -> bool {
        self.col.is_empty()
    }
}


#[derive(Debug, Deserialize, Serialize, PartialEq, Default)]
struct Col {
    #[serde(rename = "@min")]
    min: u32,
    #[serde(rename = "@max")]
    max: u32,
    #[serde(rename = "@width", skip_serializing_if = "Option::is_none")]
    width: Option<f64>,
    #[serde(rename = "@style", skip_serializing_if = "Option::is_none")]
    style: Option<u32>,
    #[serde(rename = "@bestFit", skip_serializing_if = "Option::is_none")]
    best_fit: Option<u8>,
    #[serde(rename = "@customWidth")]
    custom_width: u8,
}

impl Col {
    fn new(min: u32, max: u32, custom_width: u8, width: Option<f64>, style: Option<u32>, best_fit: Option<u8>) -> Col {
        Col {
            min,
            max,
            custom_width,
            width,
            style,
            best_fit,
        }
    }
}
