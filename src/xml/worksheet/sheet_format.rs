use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Deserialize, Serialize)]
pub(crate) struct SheetFormatPr {
    #[serde(rename = "@defaultColWidth", skip_serializing_if = "Option::is_none")]
    default_col_width: Option<f64>,
    #[serde(rename = "@defaultRowHeight")]
    default_row_height: f64,
    #[serde(rename = "@customHeight", skip_serializing_if = "Option::is_none")]
    custom_height: Option<u8>,
    #[serde(rename = "@zeroHeight", skip_serializing_if = "Option::is_none")]
    zero_height: Option<u8>,
    #[serde(rename = "@outlineLevelCol", skip_serializing_if = "Option::is_none")]
    outline_level_col: Option<u8>,
    #[serde(rename(serialize = "@x14ac:dyDescent", deserialize = "@dyDescent"), skip_serializing_if = "Option::is_none")]
    x14ac_dy_descent: Option<f64>,
}

impl Default for SheetFormatPr {
    fn default() -> SheetFormatPr {
        SheetFormatPr {
            default_col_width: None,
            default_row_height: 15.0,
            custom_height: None,
            zero_height: None,
            outline_level_col: None,
            x14ac_dy_descent: None,
        }
    }
}

impl SheetFormatPr {
    pub(crate) fn set_default_row_height(&mut self, height: f64) {
        self.custom_height = Some(1);
        self.default_row_height = height;
    }

    pub(crate) fn get_default_row_height(&self) -> f64 {
        self.default_row_height
    }

    pub(crate) fn set_default_col_width(&mut self, width: f64) {
        self.custom_height = Some(1);
        self.default_col_width = Some(width);
    }

    pub(crate) fn get_default_col_width(&self) -> f64 {
        self.default_col_width.unwrap_or(8.11)
    }

    pub(crate) fn hide_unused_rows(&mut self, hide: bool) {
        self.zero_height = Some(hide as u8);
    }

    pub(crate) fn set_outline_level_col(&mut self, col_level: u8) {
        self.outline_level_col = Some(col_level)
    }
}