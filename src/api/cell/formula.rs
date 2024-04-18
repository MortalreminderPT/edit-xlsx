use crate::api::cell::location::{Location, LocationRange};

#[derive(Clone, Debug, Default)]
pub(crate) struct Formula {
    pub(crate) formula: String,
    pub(crate) formula_type: Option<String>,
    pub(crate) si: Option<i32>,
    pub(crate) formula_ref: Option<String>,
}

impl Formula {
    pub(crate) fn new(formula_content: &str) -> Formula {
        Formula {
            formula: formula_content.to_string(),
            formula_type: None,
            si: None,
            formula_ref: None,
        }
    }

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