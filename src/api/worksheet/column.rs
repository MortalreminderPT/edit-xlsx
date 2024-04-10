#[derive(Copy, Clone, Default)]
pub struct Column {
    pub width: Option<f64>,
    pub(crate) style: Option<u32>,
    pub outline_level: Option<u32>,
    pub hidden: Option<u8>,
    pub collapsed: Option<u8>,
}