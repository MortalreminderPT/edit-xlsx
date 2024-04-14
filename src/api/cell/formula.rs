use crate::api::cell::location::{Location, LocationRange};

#[derive(Clone, Debug, Default)]
pub(crate) struct Formula {
    pub(crate) formula: String,
    pub(crate) formula_type: Option<String>,
    pub(crate) si: Option<i32>,
    pub(crate) formula_ref: Option<String>,
}

impl Formula {
    pub(crate) fn new_array_formula<L: Location>(formula_content: &str, loc: &L) -> Formula {
        Formula {
            formula: formula_content.to_string(),
            formula_type: Some("array".to_string()),
            si: None,
            formula_ref: Some(loc.to_ref()),
        }
    }

    pub(crate) fn new_array_formula_by_range<L: LocationRange>(formula_content: &str, loc_range: &L) -> Formula {
        Formula {
            formula: formula_content.to_string(),
            formula_type: Some("array".to_string()),
            si: None,
            formula_ref: Some(loc_range.to_range_ref()),
        }
    }
}

// #[derive(PartialEq, Clone, Debug)]
// pub(crate) enum FormulaType {
//     Formula(String),
//     OldFormula(String),
//     ArrayFormula(String),
//     DynamicArrayFormula(String),
// }
// 
// impl Default for FormulaType {
//     fn default() -> Self {
//         FormulaType::Formula("array".to_string())
//     }
// }

// impl FormulaType {
    // pub(crate) fn to_formula_ref(self) -> (Option<String>, Option<String>) {
    //     match self {
    //         FormulaType::Formula(formula_ref) => (Some(String::from("array")), Some(formula_ref)),
    //         FormulaType::OldFormula(formula_ref) => (Some(String::from("array")), Some(formula_ref)),
    //         FormulaType::ArrayFormula(formula_ref) => (Some(String::from("array")), Some(formula_ref)),
    //         FormulaType::DynamicArrayFormula(formula_ref) => (Some(String::from("array")), Some(formula_ref)),
    //     }
    // }
// }