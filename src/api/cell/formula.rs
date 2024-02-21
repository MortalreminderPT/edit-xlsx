pub(crate) enum FormulaType {
    Formula(String),
    ArrayFormula(String),
    DynamicArrayFormula(String),
}

impl FormulaType {
    pub(crate) fn to_formula_ref(self) -> (Option<String>, Option<String>) {
        match self {
            FormulaType::Formula(formula_ref) => (Some(String::from("array")), Some(formula_ref)),
            FormulaType::ArrayFormula(formula_ref) => (Some(String::from("array")), Some(formula_ref)),
            FormulaType::DynamicArrayFormula(formula_ref) => (Some(String::from("array")), Some(formula_ref)),
        }
    }
}