#[derive(Copy, Clone)]
pub enum FormatColor<'a> {
    Default,
    RGB(&'a str),
}

impl Default for FormatColor<'_> {
    fn default() -> Self {
        Self::Default
    }
}