use serde::{Deserialize, Serialize};

pub(crate) enum FormulaType {
    Formula,
    ArrayFormula(String),
    DynamicArrayFormula(String),
}

impl FormulaType {
    fn to_formula_ref(self) -> (Option<String>, Option<String>) {
        match self {
            FormulaType::Formula => (None, None),
            FormulaType::ArrayFormula(formula_ref) => (Some(String::from("array")), Some(formula_ref)),
            FormulaType::DynamicArrayFormula(formula_ref) => (Some(String::from("array")), Some(formula_ref)),
        }
    }
}

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
        let (formula_type, formula_ref) = formula_type.to_formula_ref();
        Formula {
            formula: formula.to_string(),
            formula_type,
            formula_ref,
        }
    }
}
