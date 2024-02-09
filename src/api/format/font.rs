use crate::FormatColor;

pub struct FormatFont<'a> {
    pub(crate) bold: bool,
    pub(crate) italic: bool,
    pub(crate) underline: bool,
    pub(crate) size: u8,
    pub(crate) color: FormatColor<'a>,
}

impl Default for FormatFont<'_> {
    fn default() -> Self {
        FormatFont {
            bold: false,
            italic: false,
            underline: false,
            size: 11,
            color: Default::default(),
        }
    }
}