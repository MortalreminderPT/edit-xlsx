use crate::FormatColor;
use crate::xml::common::FromFormat;
use crate::xml::style::color::Color;
use crate::xml::style::fill::Fill;

#[derive(Clone)]
pub struct FormatFill {
    pub(crate) pattern_type: String,
    pub(crate) fg_color: FormatColor
}

impl Default for FormatFill<> {
    fn default() -> Self {
        FormatFill {
            pattern_type: "none".to_string(),
            fg_color: Default::default(),
        }
    }
}

impl FromFormat<FormatFill> for Fill {
    fn set_attrs_by_format(&mut self, format: &FormatFill) {
        let fg_color = Color::from_format(&format.fg_color);
        self.pattern_fill.fg_color = Some(fg_color);
        self.pattern_fill.pattern_type = String::from(&format.pattern_type);
    }

    fn set_format<'a>(&self, format: &'a mut FormatFill) {
        // let mut format_new = FormatFill::default();
        format.fg_color = self.pattern_fill.fg_color.as_ref().get_format();
        format.pattern_type = self.pattern_fill.pattern_type.to_string();
        // format.fg_color = self.pattern_fill.fg_color.as_ref().unwrap().get_format();
        // format_new.pattern_type = &self.pattern_fill.pattern_type.as_str();//.to_string();
        // *format = format_new
    }
}