use crate::api::cell::formula::Formula;
use crate::api::cell::rich_text::RichText;
use crate::api::cell::values::{CellDisplay, CellType, CellValue};
use crate::Format;

pub mod formula;
pub mod location;
pub mod values;
pub mod rich_text;

#[derive(Clone, Debug, Default)]
pub struct Cell<T: CellDisplay + CellValue> {
    pub text: Option<T>,
    pub rich_text: Option<RichText>,
    pub format: Option<Format>,
    pub hyperlink: Option<String>,
    pub(crate) formula: Option<Formula>,
    pub(crate) cell_type: Option<CellType>,
    pub(crate) style: Option<u32>,
}