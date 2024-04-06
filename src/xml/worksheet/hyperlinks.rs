use serde::{Deserialize, Serialize};
use crate::api::cell::location::Location;

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
pub(crate) struct Hyperlinks {
    #[serde(rename = "hyperlink")]
    hyperlink: Vec<Hyperlink>
}

impl Hyperlinks {
    pub(crate) fn add_hyperlink<L: Location>(&mut self, loc: &L, r_id: u32) {
        let hyperlink = Hyperlink::new(&loc.to_ref(), r_id);
        self.hyperlink.push(hyperlink)
    }
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct Hyperlink {
    #[serde(rename = "@ref")]
    hyperlink_ref: String,
    #[serde(rename = "@location", skip_serializing_if = "Option::is_none")]
    location: Option<String>,
    #[serde(rename(serialize = "@r:id", deserialize = "@id"), skip_serializing_if = "Option::is_none")]
    r_id: Option<String>,
    #[serde(rename = "@display", skip_serializing_if = "Option::is_none")]
    display: Option<String>,
    #[serde(rename = "@tooltip", skip_serializing_if = "Option::is_none")]
    tooltip: Option<String>,
    #[serde(rename(serialize = "@xr:uid", deserialize = "@uid"), skip_serializing_if = "String::is_empty")]
    uid: String,
}

impl Hyperlink {
    fn new(hyperlink_ref: &str, r_id: u32) -> Self {
        Self {
            hyperlink_ref: String::from(hyperlink_ref),
            location: None,
            display: None,
            tooltip: None,
            r_id: Some(format!("rId{r_id}")),
            uid: "".to_string(),
        }
    }
}
