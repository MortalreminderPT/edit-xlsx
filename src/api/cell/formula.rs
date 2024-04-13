use crate::api::cell::formula::FormulaType::Formula;

#[derive(PartialEq, Clone, Debug)]
pub(crate) enum FormulaType {
    Formula(String),
    OldFormula(String),
    ArrayFormula(String),
    DynamicArrayFormula(String),
}

impl Default for FormulaType {
    fn default() -> Self {
        Formula("array".to_string())
    }
}

impl FormulaType {
    pub(crate) fn to_formula_ref(self) -> (Option<String>, Option<String>) {
        match self {
            FormulaType::Formula(formula_ref) => (Some(String::from("array")), Some(formula_ref)),
            FormulaType::OldFormula(formula_ref) => (Some(String::from("array")), Some(formula_ref)),
            FormulaType::ArrayFormula(formula_ref) => (Some(String::from("array")), Some(formula_ref)),
            FormulaType::DynamicArrayFormula(formula_ref) => (Some(String::from("array")), Some(formula_ref)),
        }
    }
}