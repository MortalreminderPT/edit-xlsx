use crate::xml::common::FromFormat;
use crate::xml::style::color::Color;

#[derive(Clone)]
pub enum FormatColor {
    Default,
    RGB(String),
    Index(u8),
    Theme(u8, f64)
}


impl Default for FormatColor {
    fn default() -> Self {
        Self::Default
    }
}


impl FromFormat<FormatColor> for Color {
    fn set_attrs_by_format(&mut self, format: &FormatColor) {
        match format {
            FormatColor::Default => *self = Color::default(),
            FormatColor::RGB(color) => *self = Color::from_rgb(color),
            FormatColor::Index(id) => *self = Color::from_index(*id),
            FormatColor::Theme(theme, tint) => *self = Color::from_theme(*theme, *tint),
        }
    }

    fn set_format(&self, format: &mut FormatColor) {
        *format = if let Some(id) = self.indexed {
            FormatColor::Index(id)
        } else if let (Some(theme), tint) = (self.theme, self.tint) {
            FormatColor::Theme(theme, tint.unwrap_or_default())
        } else if let Some(color) = &self.rgb {
            FormatColor::RGB(color.clone())
        } else {
            FormatColor::Default
        };
    }
}