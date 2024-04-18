use crate::xml::common::FromFormat;
use crate::xml::style::color::Color;

#[derive(Clone, Debug, PartialEq)]
pub enum FormatColor {
    Default,
    // RGB(String),
    Index(u8),
    Theme(u8, f64),
    RGB(u8, u8, u8),
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
            FormatColor::RGB(r, g, b) => *self = Color::from_rgb(*r, *g, *b),
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
            let argb: Vec<u8> = color
                .chars()
                .collect::<Vec<char>>()
                .chunks(2)
                .map(|chunk| {
                    let hex_string: String = chunk.iter().collect();
                    u8::from_str_radix(&hex_string, 16).unwrap()
                })
                .collect();
            FormatColor::RGB(argb[1], argb[2], argb[3])
        } else {
            FormatColor::Default
        };
    }
}