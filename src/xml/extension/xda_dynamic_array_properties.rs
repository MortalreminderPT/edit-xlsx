use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct XdaDynamicArrayProperties {
    #[serde(rename = "@fDynamic")]
    f_dynamic: u8,
    #[serde(rename = "@fCollapsed")]
    f_collapsed: u8,
}

impl Default for XdaDynamicArrayProperties {
    fn default() -> Self {
        Self {
            f_dynamic: 1,
            f_collapsed: 0,
        }
    }
}