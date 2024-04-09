#[derive(Copy, Clone)]
pub enum FormatColor<'a> {
    Default,
    RGB(&'a str),
    Index(u8),
    Theme(u8, f64)
}

impl Default for FormatColor<'_> {
    fn default() -> Self {
        Self::Default
    }
}