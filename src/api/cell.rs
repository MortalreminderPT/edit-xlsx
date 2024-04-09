use crate::api::cell::values::{CellDisplay, CellValue};
use crate::Format;

pub mod formula;
pub mod location;
pub mod values;

pub(crate) struct Cell<'a, Cell: CellDisplay + CellValue> {
    pub(crate) text: Option<Cell>,
    pub(crate) format: Option<Format<'a>>,
    pub(crate) hyperlink: Option<String>,
    pub(crate) formula: Option<String>,
}