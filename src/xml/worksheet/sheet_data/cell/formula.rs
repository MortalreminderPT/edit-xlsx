use serde::{Deserialize, Serialize};
use crate::api::cell::formula::FormulaType;

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Formula {
    #[serde(rename = "$value")]
    formula: String,
    #[serde(rename = "@t", skip_serializing_if = "Option::is_none")]
    formula_type: Option<String>,
    #[serde(rename = "@ref", skip_serializing_if = "Option::is_none")]
    formula_ref: Option<String>,
}

impl Formula {
    pub(crate) fn from_formula_type(formula: &str, formula_type: FormulaType) -> Formula {
        // let mut formula = formula.trim_matches(|f| f == '{' || f == '}').to_string();
        // if formula.starts_with("=") {
        //     formula = formula.split_off(1);
        //     formula = format!("_xlfn._xlws.{}", formula);
        // }
        let (formula_type, formula_ref) = formula_type.to_formula_ref();
        Formula {
            formula: formula.to_string(),
            formula_type,
            formula_ref,
        }
    }
}
