#[derive(Copy, Clone, Default, Debug)]
pub struct Column {
    pub(crate) width: Option<f64>,
    pub(crate) style: Option<u32>,
    pub(crate) outline_level: Option<u32>,
    pub(crate) hidden: Option<u8>,
    pub(crate) collapsed: Option<u8>,
}

impl Column {
    pub fn set_width(mut self, width: f64) -> Self {
        self.width = Some(width);
        self
    }
    pub fn set_outline_level(mut self, outline_level: u32) -> Self {
        self.outline_level = Some(outline_level);
        self
    }
    pub fn hide(mut self) -> Self {
        self.hidden = Some(1);
        self
    }
    pub fn collapse(mut self) -> Self {
        self.collapsed = Some(1);
        self
    }
}