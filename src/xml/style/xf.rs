use serde::{Deserialize, Serialize};
use crate::xml::style::alignment::Alignment;
use crate::xml::common;

#[derive(Debug, Deserialize, Serialize, PartialEq, Clone)]
pub(crate) struct Xf {
    #[serde(rename = "@numFmtId", default)]
    num_fmt_id: u32,
    #[serde(rename = "@fontId", default)]
    pub(crate) font_id: u32,
    #[serde(rename = "@fillId", default)]
    pub(crate) fill_id: u32,
    #[serde(rename = "@borderId", default)]
    pub(crate) border_id: u32,
    #[serde(rename = "@xfId", default, skip_serializing_if = "common::is_zero")]
    xf_id: u32,
    #[serde(rename = "@applyFont", default, skip_serializing_if = "common::is_zero")]
    apply_font: u32,
    #[serde(rename = "@applyAlignment", default, skip_serializing_if = "common::is_zero")]
    pub(crate) apply_alignment: u32,
    #[serde(rename = "alignment", skip_serializing_if = "Option::is_none")]
    pub(crate) alignment: Option<Alignment>,
}

impl Xf {
    pub(crate) fn default() -> Xf {
        Xf {
            num_fmt_id: 0,
            font_id: 0,
            fill_id: 0,
            border_id: 0,
            xf_id: 0,
            apply_font: 0,
            apply_alignment: 0,
            alignment: None,
        }
    }
}
