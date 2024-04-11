use serde::{Deserialize, Serialize};
use crate::api::format::FormatFill;
use crate::xml::common::FromFormat;
use crate::xml::style::color::Color;


#[derive(Debug, Clone, Deserialize, Serialize, Default)]
pub(crate) struct Fills {
    #[serde(rename = "@count", default)]
    count: u32,
    #[serde(rename = "fill", default)]
    fills: Vec<Fill>
}

impl Fills {
    pub(crate) fn add_fill(&mut self, fill: &Fill) -> u32 {
        for i in 0..self.fills.len() {
            if self.fills[i] == *fill {
                return i as u32;
            }
        }
        self.count += 1;
        self.fills.push(fill.clone());
        self.fills.len() as u32 - 1
    }
    
    pub(crate) fn get_fill(&self, id: u32) -> Option<&Fill> {
        self.fills.get(id as usize)
    }
}

#[derive(Debug, Deserialize, Serialize, Clone, PartialEq, Default)]
pub(crate) struct Fill {
    #[serde(rename = "patternFill")]
    pub(crate) pattern_fill: PatternFill
}

#[derive(Debug, Deserialize, Serialize, Clone, PartialEq)]
pub(crate) struct PatternFill {
    #[serde(rename = "@patternType")]
    pub(crate) pattern_type: String,
    #[serde(rename = "fgColor", default, skip_serializing_if = "Color::is_empty")]
    pub(crate) fg_color: Color,
    #[serde(rename = "bgColor", default, skip_serializing_if = "Color::is_empty")]
    pub(crate) bg_color: Color,
}

impl Default for PatternFill {
    fn default() -> PatternFill {
        PatternFill {
            pattern_type: "none".to_string(),
            fg_color: Color::default(),
            bg_color: Color::default(),
        }
    }
}