use serde::{Deserialize, Serialize};
use crate::Filters as ApiFilters;
use crate::Filter as ApiFilter;

#[derive(Debug, Deserialize, Serialize, Default)]
pub(crate) struct AutoFilter {
    #[serde(rename = "@ref")]
    pub(crate) sqref: String,
    #[serde(rename = "filterColumn")]
    filter_column: Vec<FilterColumn>,
}

impl AutoFilter {
    pub(crate) fn add_filters(&mut self, col: u32, filters: &ApiFilters) {
        let mut filter_column = FilterColumn::new(col, Some(filters.and));
        filter_column.filters.add_api_filters(filters);
        self.filter_column.push(filter_column);
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct FilterColumn {
    #[serde(rename = "@colId")]
    col_id: u32,
    #[serde(rename = "@and", skip_serializing_if = "Option::is_none")]
    and: Option<u8>,
    #[serde(rename = "customFilters")]
    filters: Filters
}

impl FilterColumn {
    fn new(col: u32, and: Option<u8>) -> FilterColumn {
        FilterColumn {
            col_id: col,
            and,
            filters: Filters::default(),
        }
    }
}

#[derive(Debug, Deserialize, Serialize, Default)]
struct Filters {
    #[serde(rename = "customFilter")]
    filters: Vec<Filter>
}

impl Filters {
    fn add_api_filters(&mut self, api_filters: &ApiFilters) {
        // self.and = Some(api_filters.and);
        api_filters.filters.iter().for_each(|f|{
            self.filters.push(
                Filter::from_api_filter(f)
            );
        });
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct Filter {
    #[serde(rename = "@val")]
    val: Option<String>,
    #[serde(rename = "@operator")]
    operator: Option<String>
}

impl Filter {
    fn from_val(val: &str) -> Filter {
        Filter {
            val: Some(String::from(val)),
            operator: None,
        }
    }

    fn from_api_filter(api_filter: &ApiFilter) -> Filter {
        Filter {
            val: Some(api_filter.val.to_string()),
            operator: if let Some(operator) = api_filter.operator { Some(operator.to_string()) } else { None },
        }
    }
}