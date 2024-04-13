use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Deserialize, Serialize)]
pub(crate) struct X15WorkbookPr {
    #[serde(rename = "@chartTrackingRefBase")]
    chart_tracking_ref_base: u32
}

impl Default for X15WorkbookPr {
    fn default() -> Self {
        X15WorkbookPr {
            chart_tracking_ref_base: 1,
        }
    }
}