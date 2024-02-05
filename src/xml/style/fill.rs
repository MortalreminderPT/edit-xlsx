use serde::{Deserialize, Serialize};
use crate::xml::common::Color;

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