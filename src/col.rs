use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct Col {
    #[serde(rename = "@min")]
    min: u32,
    #[serde(rename = "@max")]
    max: u32,
    #[serde(rename = "@customWidth")]
    custom_width: bool,
    #[serde(rename = "@width")]
    width: f64,
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct Cols {
    col: Vec<Col>
}