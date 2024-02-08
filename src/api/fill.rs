use crate::FormatColor;

pub(crate) struct FormatFill<'a> {
    fg_color: FormatColor<'a>
}

impl Default for FormatFill {
    fn default() -> Self {
        FormatFill {
            fg_color: Default::default(),
        }
    }
}