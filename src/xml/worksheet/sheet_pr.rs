use serde::{Deserialize, Serialize};
use crate::FormatColor;
use crate::xml::common::FromFormat;
use crate::xml::style::color::Color;

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
pub(crate) struct SheetPr {
    #[serde(rename = "tabColor", skip_serializing_if = "Option::is_none")]
    tab_color: Option<Color>,
    #[serde(rename = "outlinePr", skip_serializing_if = "Option::is_none")]
    outline_pr: Option<OutlinePr>,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct OutlinePr {
    #[serde(rename = "@applyStyles")]
    apply_styles: u8,
    #[serde(rename = "@summaryBelow")]
    summary_below: u8,
    #[serde(rename = "@summaryRight")]
    summary_right: u8,
    #[serde(rename = "@showOutlineSymbols")]
    show_outline_symbols: u8,
}

impl Default for OutlinePr {
    fn default() -> Self {
        Self {
            apply_styles: 0,
            summary_below: 1,
            summary_right: 1,
            show_outline_symbols: 1,
        }
    }
}

impl SheetPr {
    fn new(color: &FormatColor) -> SheetPr {
        SheetPr {
            tab_color: Default::default(), //Color::from_format(color),
            outline_pr: None,
        }
    }

    pub(crate) fn set_tab_color(&mut self, color: &FormatColor) {
        let tab_color = Color::from_format(color);
        self.tab_color = Some(tab_color);
    }

    pub(crate) fn set_outline_pr(&mut self, visible: bool, symbols_below: bool, symbols_right: bool, auto_style: bool) {
        let mut outline_pr = self.outline_pr.take().unwrap_or_default();
        outline_pr.apply_styles = auto_style as u8;
        outline_pr.show_outline_symbols = visible as u8;
        outline_pr.summary_below = symbols_below as u8;
        outline_pr.summary_right = symbols_right as u8;
        self.outline_pr = Some(outline_pr);
    }
}
