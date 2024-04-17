use crate::FormatColor;
use crate::xml::common::FromFormat;
use crate::xml::style::color::Color;
use crate::xml::style::fill::{Fill, PatternFill};

#[derive(Clone, Debug)]
pub struct FormatFill {
    pub pattern_type: String,
    pub fg_color: FormatColor,
    pub bg_color: FormatColor
}

impl Default for FormatFill<> {
    fn default() -> Self {
        FormatFill {
            pattern_type: "none".to_string(),
            fg_color: FormatColor::default(),
            bg_color: FormatColor::Index(64),
        }
    }
}

impl FromFormat<FormatFill> for Fill {
    fn set_attrs_by_format(&mut self, format: &FormatFill) {
        self.pattern_fill.fg_color = Color::from_format(&format.fg_color);
        self.pattern_fill.bg_color = Color::from_format(&format.bg_color);
        self.pattern_fill.pattern_type = String::from(&format.pattern_type);
    }

    fn set_format(&self, format: &mut FormatFill) {
        format.fg_color = self.pattern_fill.fg_color.get_format();
        format.bg_color = self.pattern_fill.bg_color.get_format();
        format.pattern_type = self.pattern_fill.pattern_type.to_string();
    }
}