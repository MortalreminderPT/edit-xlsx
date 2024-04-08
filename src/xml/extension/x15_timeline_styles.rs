use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
pub(crate) struct X15TimelineStyles {
    #[serde(rename = "@defaultTimelineStyle", skip_serializing_if = "String::is_empty")]
    default_timeline_style: String
}
