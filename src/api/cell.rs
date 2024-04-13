use crate::api::cell::formula::FormulaType;
use crate::api::cell::values::{CellDisplay, CellType, CellValue};
use crate::Format;

pub mod formula;
pub mod location;
pub mod values;

#[derive(Clone, Debug, Default)]
pub struct Cell<T: CellDisplay + CellValue> {
    pub text: Option<T>,
    pub format: Option<Format>,
    pub hyperlink: Option<String>,
    pub formula: Option<String>,
    pub(crate) cell_type: Option<CellType>,
    pub(crate) formula_type: Option<FormulaType>,
    pub(crate) formula_ref: Option<String>,
    pub(crate) style: Option<u32>,
}
//
// impl<T: CellDisplay + CellValue> Default for Cell<T> {
//     fn default() -> Self {
//         Self {
//             text: None,
//             format: None,
//             hyperlink: None,
//             formula: None,
//             formula_type: None,
//             style: None,
//         }
//     }
// }