use serde::{Deserialize, Serialize};
use crate::xml::style::color::Color;


#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Fills {
    #[serde(rename = "@count", default)]
    count: u32,
    #[serde(rename = "fill", default)]
    fills: Vec<Fill>
}

impl Fills {
    fn default() -> Fills {
        Fills {
            count: 0,
            fills: vec![],
        }
    }

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
}

#[derive(Debug, Deserialize, Serialize, Clone, PartialEq)]
pub(crate) struct Fill {
    #[serde(rename = "patternFill")]
    pub(crate) pattern_fill: PatternFill
}

impl Fill {
    pub(crate) fn default() -> Fill {
        Fill {
            pattern_fill: PatternFill::default(),
        }
    }
}

#[derive(Debug, Deserialize, Serialize, Clone, PartialEq)]
pub(crate) struct PatternFill {
    #[serde(rename = "@patternType")]
    pub(crate) pattern_type: String,
    #[serde(rename = "fgColor", skip_serializing_if = "Option::is_none")]
    pub(crate) fg_color: Option<Color>,
    #[serde(rename = "bgColor", skip_serializing_if = "Option::is_none")]
    pub(crate) bg_color: Option<Color>,
}

impl PatternFill {
    fn default() -> PatternFill {
        PatternFill {
            pattern_type: "none".to_string(),
            fg_color: None,
            bg_color: None,
        }
    }
}