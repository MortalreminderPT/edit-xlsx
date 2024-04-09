use serde::{Deserialize, Serialize};
use crate::api::format::FormatFill;
use crate::xml::common::FromFormat;
use crate::xml::style::color::Color;


#[derive(Debug, Clone, Deserialize, Serialize)]
pub(crate) struct Fills {
    #[serde(rename = "@count", default)]
    count: u32,
    #[serde(rename = "fill", default)]
    fills: Vec<Fill>
}

impl Default for Fills {
    fn default() -> Fills {
        Fills {
            count: 0,
            fills: vec![],
        }
    }
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

#[derive(Debug, Deserialize, Serialize, Clone, PartialEq)]
pub(crate) struct Fill {
    #[serde(rename = "patternFill")]
    pub(crate) pattern_fill: PatternFill
}

// impl Fill {
//     pub(crate) fn default() -> Fill {
//         Fill {
//             pattern_fill: PatternFill::default(),
//         }
//     }
// }

impl Default for Fill {
    fn default() -> Self {
        Self {
            pattern_fill: Default::default(),
        }
    }
}

impl FromFormat<FormatFill<'_>> for Fill {
    fn set_attrs_by_format(&mut self, format: &FormatFill) {
        let fg_color = Color::from_format(&format.fg_color);
        self.pattern_fill.fg_color = Some(fg_color);
        self.pattern_fill.pattern_type = String::from(format.pattern_type);
    }

    fn set_format<'a>(&self, format: &'a mut FormatFill) {
        let mut format_new = FormatFill::default();
        format_new.fg_color = self.pattern_fill.fg_color.as_ref().get_format();
        // format.fg_color = self.pattern_fill.fg_color.as_ref().unwrap().get_format();
        // format_new.pattern_type = &self.pattern_fill.pattern_type.as_str();//.to_string();
        *format = format_new
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

impl Default for PatternFill {
    fn default() -> PatternFill {
        PatternFill {
            pattern_type: "none".to_string(),
            fg_color: None,
            bg_color: None,
        }
    }
}