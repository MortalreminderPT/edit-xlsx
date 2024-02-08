pub enum FormatColor<'a> {
    RGB(&'a str),
}

impl Default for FormatColor {
    fn default() -> Self {
        Self::RGB("")
    }
}