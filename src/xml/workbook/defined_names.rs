use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize, Default)]
pub(crate) struct DefinedNames {
    #[serde(rename = "definedName", skip_serializing_if = "Vec::is_empty")]
    defined_names: Vec<DefinedName>
}

impl DefinedNames {
    pub(crate) fn is_empty(&self) -> bool {
        self.defined_names.is_empty()
    }

    pub(crate) fn add_define_name(&mut self, name: &str, value: &str, local_sheet_id: Option<u32>) {
        let defined_name = DefinedName::new(name, value, local_sheet_id);
        self.defined_names.push(defined_name)
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