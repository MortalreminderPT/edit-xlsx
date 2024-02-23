use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct SheetFormatPr {
    #[serde(rename = "@defaultRowHeight")]
    default_row_height: f64,
    #[serde(rename = "@customHeight", skip_serializing_if = "Option::is_none")]
    custom_height: Option<u8>,
    #[serde(rename = "@zeroHeight", skip_serializing_if = "Option::is_none")]
    zero_height: Option<u8>,
    #[serde(rename = "@outlineLevelCol", skip_serializing_if = "Option::is_none")]
    outline_level_col: Option<u8>,
}

impl Default for SheetFormatPr {
    fn default() -> SheetFormatPr {
        SheetFormatPr {
            default_row_height: 15.0,
            custom_height: None,
            zero_height: None,
            outline_level_col: None,
        }
    }
}

impl SheetFormatPr {
    pub(crate) fn set_default_row_height(&mut self, height: f64) {
        self.custom_height = Some(1);
        self.default_row_height = height;
    }

    pub(crate) fn hide_unused_rows(&mut self, hide: bool) {
        self.zero_height = Some(hide as u8);
    }

    pub(crate) fn set_outline_level_col(&mut self, col_level: u8) {
        self.outline_level_col = Some(col_level)
    }
}