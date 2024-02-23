use std::cmp::max_by;
use serde::{Deserialize, Serialize};
use crate::result::ColResult;

#[derive(Debug, Deserialize, Serialize, PartialEq, Default)]
pub(crate) struct Cols {
    col: Vec<Col>
}

impl Cols {
    pub(crate) fn update_col(&mut self, min: u32, max: u32, width: Option<f64>, style: Option<u32>, hidden: Option<u8>, best_fit: Option<u8>) -> ColResult<()> {
        let col = self.get_or_new(min, max);
        col.update_width(width);
        col.style = style;
        col.hidden = hidden;
        col.best_fit = best_fit;
        // if let Some(width) = width {
        //     col.update_width(width);
        // }
        // if let Some(style) = style {
        //     col.style = Some(style);
        // }
        // if let Some(hidden) = hidden {
        //     col.hidden = Some(hidden);
        // }
        // if let Some(best_fit) = best_fit {
        //     col.best_fit = Some(best_fit);
        // }
        Ok(())
    }

    fn get_or_new(&mut self, min: u32, max: u32) -> &mut Col {
        let len = self.col.len();
        for i in 0..len {
            if self.col[i].min == min && self.col[i].max == max {
                return &mut self.col[i];
            }
        }
        self.col.push(Col::new(min, max));
        &mut self.col[len]
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
    #[serde(rename = "@hidden", skip_serializing_if = "Option::is_none")]
    hidden: Option<u8>,
    #[serde(rename = "@customWidth")]
    custom_width: u8,
}

impl Col {
    fn new(min: u32, max: u32) -> Col {
        Col {
            min,
            max,
            width: None,
            style: None,
            best_fit: None,
            hidden: None,
            custom_width: 0,
        }
    }

    fn update_width(&mut self, width: Option<f64>) {
        self.width = width;
        if let Some(_) = width {
            self.custom_width = 1;
        }
    }

    // fn new(min: u32, max: u32, custom_width: u8, width: Option<f64>, style: Option<u32>, best_fit: Option<u8>, hidden: Option<u8>) -> Col {
    //     Col {
    //         min,
    //         max,
    //         custom_width,
    //         width,
    //         style,
    //         best_fit,
    //         hidden,
    //     }
    // }
}
