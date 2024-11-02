use serde::{Deserialize, Serialize};
use crate::Filters as ApiFilters;
use crate::Filter as ApiFilter;
use crate::xml::common::XmlnsAttrs;

#[derive(Debug, Clone, Deserialize, Serialize)]
pub(crate) struct AutoFilter {
    #[serde(flatten)]
    pub(crate) xmlns_attrs: XmlnsAttrs,
    #[serde(rename = "@ref", default, skip_serializing_if = "String::is_empty")]
    pub(crate) sqref: String,
    #[serde(rename = "filterColumn", default, skip_serializing_if = "Vec::is_empty")]
    filter_column: Vec<FilterColumn>
}

impl Default for AutoFilter {
    fn default() -> Self {
        AutoFilter {
            xmlns_attrs: XmlnsAttrs::default_none(),
            sqref: "".to_string(),
            filter_column: vec![],
        }
    }
}

impl AutoFilter {
    pub(crate) fn add_filters(&mut self, col: u32, filters: &ApiFilters) {
        let mut filter_column = FilterColumn::new(col);
        if filters.is_custom_filters() {
            filter_column.custom_filters.get_or_insert(Default::default()).add_custom_filters(filters);
        } else {
            filter_column.filters.get_or_insert(Default::default()).add_filters(filters);
        }
        self.filter_column.push(filter_column);
    }
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct FilterColumn {
    #[serde(rename = "@colId")]
    col_id: u32,
    #[serde(rename = "filters", skip_serializing_if = "Option::is_none")]
    filters: Option<Filters>,
    #[serde(rename = "customFilters", skip_serializing_if = "Option::is_none")]
    custom_filters: Option<Filters>,
}

impl FilterColumn {
    fn new(col: u32) -> FilterColumn {
        FilterColumn {
            col_id: col - 1,
            filters: Default::default(),
            custom_filters: Default::default(),
        }
    }
}

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
struct Filters {
    #[serde(rename = "@and", skip_serializing_if = "Option::is_none")]
    and: Option<u8>,
    #[serde(rename = "@blank", skip_serializing_if = "Option::is_none")]
    blank: Option<u8>,
    #[serde(rename = "filter")]
    filters: Vec<Filter>,
    #[serde(rename = "customFilter")]
    custom_filters: Vec<Filter>
}

impl Filters {
    fn add_custom_filters(&mut self, api_filters: &ApiFilters) {
        self.and = api_filters.and;
        api_filters.filters.iter().for_each(|f| {
            self.custom_filters.push(
                Filter::from_api_filter(f)
            );
        });
    }

    fn add_filters(&mut self, api_filters: &ApiFilters) {
        self.and = api_filters.and;
        self.blank = api_filters.blank;
        api_filters.filters.iter().for_each(|f| {
            self.filters.push(
                Filter::from_api_filter(f)
            );
        });
    }
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct Filter {
    #[serde(rename = "@val")]
    val: Option<String>,
    #[serde(rename = "@operator", skip_serializing_if = "Option::is_none")]
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