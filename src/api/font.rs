use crate::FormatColor;

pub struct FormatFont<'a> {
    bold: bool,
    italic: bool,
    underline: bool,
    size: bool,
    color: FormatColor<'a>,
}

impl Default for FormatFont {
    fn default() -> Self {
        FormatFont {
            bold: false,
            italic: false,
            underline: false,
            size: false,
            color: Default::default(),
        }
    }
}