use serde::{Deserialize, Serialize};
use crate::api::cell::location::Location;
use crate::utils::col_helper::to_ref;

#[derive(Debug, Deserialize, Serialize, Default)]
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

#[derive(Debug, Deserialize, Serialize)]
struct Hyperlink {
    #[serde(rename = "@ref")]
    hyperlink_ref: String,
    #[serde(rename(serialize = "@r:id", deserialize = "@id"))]
    r_id: String,
    // #[serde(rename(serialize = "@xr:uid", deserialize = "@uid"))]
    // uid: String,
}

impl Hyperlink {
    fn new(hyperlink_ref: &str, r_id: u32) -> Self {
        Self {
            hyperlink_ref: String::from(hyperlink_ref),
            r_id: format!("rId{r_id}"),
        }
    }
}
