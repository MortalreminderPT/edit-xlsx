use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize, Default)]
pub(crate) struct DefinedNames {
    #[serde(rename = "definedName", skip_serializing_if = "Vec::is_empty")]
    defined_names: Vec<DefinedName>,
}

impl DefinedNames {
    pub(crate) fn is_empty(&self) -> bool {
        self.defined_names.is_empty()
    }
    /// Attempt to add the defined name.  If no sheet id provided, this is added globally
    pub(crate) fn add_define_name(&mut self, name: &str, value: &str, local_sheet_id: Option<u32>) {
        let defined_name = DefinedName::new(name, value, local_sheet_id);
        self.defined_names.push(defined_name)
    }
    /// Attempt to find the defined name.  If no sheet id provided, the global names are checked
    pub(crate) fn get_defined_name(&self, name: &str, local_sheet_id: Option<u32>) -> Option<&str> {
        self.defined_names.iter().find_map(|entry| {
            (entry.name == name && entry.local_sheet_id == local_sheet_id)
                .then_some(entry.value.as_str())
        })
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct DefinedName {
    #[serde(rename = "@name")]
    name: String,
    #[serde(rename = "@localSheetId", skip_serializing_if = "Option::is_none")]
    local_sheet_id: Option<u32>,
    #[serde(rename = "$value", default, skip_serializing_if = "String::is_empty")]
    value: String,
}

impl DefinedName {
    fn new(name: &str, value: &str, local_sheet_id: Option<u32>) -> DefinedName {
        DefinedName {
            name: String::from(name),
            local_sheet_id,
            value: String::from(value),
        }
    }
}
