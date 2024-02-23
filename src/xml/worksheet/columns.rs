use serde::{Deserialize, Serialize};
use crate::result::ColResult;

#[derive(Debug, Deserialize, Serialize, PartialEq, Default)]
pub(crate) struct Cols {
    col: Vec<Col>
}

impl Cols {
    pub(crate) fn update_col(&mut self, min: u32, max: u32, width: Option<f64>, style: Option<u32>, hidden: Option<u8>, best_fit: Option<u8>) -> ColResult<()> {
        let col = self.get_or_new_col(min, max);
        col.update_width(width);
        col.style = style;
        col.hidden = hidden;
        col.best_fit = best_fit;
        Ok(())
    }

    pub(crate) fn get_or_new_col(&mut self, min: u32, max: u32) -> &mut Col {
        let len = self.col.len();
        for i in 0..len {
            if self.col[i].min == min && self.col[i].max == max {
                return &mut self.col[i];
            }
        }
        self.col.push(Col::new(min, max));
        self.col.last_mut().unwrap()
    }
    
    pub(crate) fn get_default_style(&self, col: u32) -> Option<u32> {
        let col = self.col.iter().filter(|c| c.min <= col && c.max >= col).last();
        if let Some(col) = col {
            return col.style;
        }
        None
    }
    
    pub(crate) fn is_empty(&self) -> bool {
        self.col.is_empty()
    }
}


#[derive(Debug, Deserialize, Serialize, PartialEq, Default)]
pub(crate) struct Col {
    #[serde(rename = "@min")]
    min: u32,
    #[serde(rename = "@max")]
    max: u32,
    #[serde(rename = "@width", skip_serializing_if = "Option::is_none")]
    pub(crate) width: Option<f64>,
    #[serde(rename = "@style", skip_serializing_if = "Option::is_none")]
    pub(crate) style: Option<u32>,
    #[serde(rename = "@bestFit", skip_serializing_if = "Option::is_none")]
    pub(crate) best_fit: Option<u8>,
    #[serde(rename = "@hidden", skip_serializing_if = "Option::is_none")]
    pub(crate) hidden: Option<u8>,
    #[serde(rename = "@outlineLevel", skip_serializing_if = "Option::is_none")]
    pub(crate) outline_level: Option<u32>,
    #[serde(rename = "@collapsed", skip_serializing_if = "Option::is_none")]
    pub(crate) collapsed: Option<u8>,
    #[serde(rename = "@customWidth", skip_serializing_if = "Option::is_none")]
    pub(crate) custom_width: Option<u8>,
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
            outline_level: None,
            collapsed: None,
            custom_width: None,
        }
    }

    fn update_width(&mut self, width: Option<f64>) {
        self.width = width;
        if let Some(_) = width {
            self.custom_width = Some(1);
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
