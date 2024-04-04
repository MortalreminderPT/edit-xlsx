use crate::FormatColor;

pub struct FormatFont<'a> {
    pub(crate) bold: bool,
    pub(crate) italic: bool,
    pub(crate) underline: bool,
    pub(crate) size: f64,
    pub(crate) color: FormatColor<'a>,
    pub(crate) name: &'a str,
}

impl Default for FormatFont<'_> {
    fn default() -> Self {
        FormatFont {
            bold: false,
            italic: false,
            underline: false,
            size: 11.0,
            color: Default::default(),
            name: "Calibri",
        }
    }
}