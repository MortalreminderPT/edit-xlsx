use crate::FormatColor;

#[derive(Clone)]
pub struct FormatFill<'a> {
    pub(crate) pattern_type: &'a str,
    pub(crate) fg_color: FormatColor<'a>
}

impl Default for FormatFill<'_> {
    fn default() -> Self {
        FormatFill {
            pattern_type: "none",
            fg_color: Default::default(),
        }
    }
}