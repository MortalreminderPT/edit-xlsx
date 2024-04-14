use serde::{Deserialize, Serialize};
use crate::api::cell::formula::Formula as ApiFormula;

#[derive(Debug, Clone, Default, Deserialize, Serialize)]
pub(crate) struct Formula {
    #[serde(rename = "@t", skip_serializing_if = "Option::is_none")]
    pub(crate) formula_type: Option<String>,
    #[serde(rename = "@ref", skip_serializing_if = "Option::is_none")]
    pub(crate) formula_ref: Option<String>,
    #[serde(rename = "@si", skip_serializing_if = "Option::is_none")]
    si: Option<i32>,
    #[serde(rename = "$value", default, skip_serializing_if = "String::is_empty")]
    pub(crate) formula: String,
}

impl Formula {
    pub(crate) fn to_api_formula(&self) -> ApiFormula {
        ApiFormula {
            formula: self.formula.to_string(),
            formula_type: self.formula_type.clone(),
            si: self.si.clone(),
            formula_ref: self.formula_ref.clone(),
        }
    }

    pub(crate) fn from_api_formula(api_formula: &ApiFormula) -> Formula {
        let mut formula = Formula::default();
        if ! &api_formula.formula.is_empty() {
            formula.formula = api_formula.formula.clone();
        }
        if api_formula.formula_ref.is_some() {
            formula.formula_ref = api_formula.formula_ref.clone();
        }
        if api_formula.formula_type.is_some() {
            formula.formula_type = api_formula.formula_type.clone();
        }
        if api_formula.si.is_some() {
            formula.si = api_formula.si.clone()
        }
        formula
    }
}

impl Formula {
    // pub(crate) fn from_formula_type(formula: &str, formula_type: FormulaType) -> Formula {
    //     // let mut formula = formula.trim_matches(|f| f == '{' || f == '}').to_string();
    //     // if formula.starts_with("=") {
    //     //     formula = formula.split_off(1);
    //     //     formula = format!("_xlfn._xlws.{}", formula);
    //     // }
    //     let (formula_type, formula_ref) = formula_type.to_formula_ref();
    //     Formula {
    //         formula: formula.to_string(),
    //         formula_type,
    //         si: None,
    //         formula_ref,
    //     }
    // }

    // pub(crate) fn from_api_formula(formula: &str, formula_type: &Option<String>, formula_ref: &Option<String>) -> Formula {
    //     // let mut formula = formula.trim_matches(|f| f == '{' || f == '}').to_string();
    //     // if formula.starts_with("=") {
    //     //     formula = formula.split_off(1);
    //     //     formula = format!("_xlfn._xlws.{}", formula);
    //     // }
    //     Formula {
    //         formula: formula.to_string(),
    //         formula_type: formula_type.clone(),
    //         si: None,
    //         formula_ref: formula_ref.clone(),
    //     }
    // }

    // pub(crate) fn get_formula(&self) -> &String {
    //     &self.formula
    // }
}
