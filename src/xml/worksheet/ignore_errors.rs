use std::collections::HashMap;
use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize, Default)]
pub(crate) struct IgnoredErrors {
    #[serde(rename = "ignoredError")]
    ignored_error: Vec<IgnoredError>,
}

impl IgnoredErrors {
    pub(crate) fn from_map(error_map: HashMap<&str, String>) -> IgnoredErrors {
        let mut ignored_errors = IgnoredErrors::default();
        error_map.iter().for_each(|(error_type, sqref)|{
            ignored_errors.ignored_error.push(IgnoredError::from_sqref(error_type, sqref));
        });
        ignored_errors
    }
}

#[derive(Debug, Deserialize, Serialize, Default)]
struct IgnoredError {
    #[serde(rename = "@sqref")]
    sqref: String,
    #[serde(rename = "@numberStoredAsText")]
    number_stored_as_text: Option<u8>,
    #[serde(rename = "@evalError")]
    eval_error: Option<u8>,
    #[serde(rename = "@formulaDiffers")]
    formula_differs: Option<u8>,
    #[serde(rename = "@formulaRange")]
    formula_range: Option<u8>,
    #[serde(rename = "@formulaUnlocked")]
    formula_unlocked: Option<u8>,
    #[serde(rename = "@emptyCellReference")]
    empty_cell_reference: Option<u8>,
    #[serde(rename = "@listDataValidation")]
    list_data_validation: Option<u8>,
    #[serde(rename = "@calculatedColumn")]
    calculated_column: Option<u8>,
    #[serde(rename = "@twoDigitTextYear")]
    two_digit_text_year: Option<u8>,
}

impl IgnoredError {
    fn from_sqref(error_type: &str, sqref: &str) -> IgnoredError {
        let mut ignored_error = IgnoredError::default();
        match error_type {
            "number_stored_as_text" => {
                ignored_error.number_stored_as_text = Some(1);
                ignored_error.sqref = String::from(sqref);
            },
            "eval_error" => {
                ignored_error.eval_error = Some(1);
                ignored_error.sqref = String::from(sqref);
            },
            "formula_differs" => {
                ignored_error.formula_differs = Some(1);
                ignored_error.sqref = String::from(sqref);
            },
            "formula_range" => {
                ignored_error.formula_range = Some(1);
                ignored_error.sqref = String::from(sqref);
            },
            "formula_unlocked" => {
                ignored_error.formula_unlocked = Some(1);
                ignored_error.sqref = String::from(sqref);
            },
            "empty_cell_reference" => {
                ignored_error.empty_cell_reference = Some(1);
                ignored_error.sqref = String::from(sqref);
            },
            "list_data_validation" => {
                ignored_error.list_data_validation = Some(1);
                ignored_error.sqref = String::from(sqref);
            },
            "calculated_column" => {
                ignored_error.calculated_column = Some(1);
                ignored_error.sqref = String::from(sqref);
            },
            "two_digit_text_year" => {
                ignored_error.two_digit_text_year = Some(1);
                ignored_error.sqref = String::from(sqref);
            },
            &_ => {}
        }
        ignored_error
    }
}