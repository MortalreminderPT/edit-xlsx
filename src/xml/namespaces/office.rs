use serde::{Deserialize, Serialize};

#[derive(Debug, Default, Deserialize, Serialize)]
// #[serde(rename(serialize = "o:shapelayout", deserialize = "shapelayout"))]
pub(crate) struct ShapeLayout {
    #[serde(rename(serialize = "@v:ext", deserialize = "@ext"))]
    v_ext: String,
    #[serde(rename(serialize = "o:idmap", deserialize = "idmap"))]
    id_map: IdMap,
}

#[derive(Debug, Default, Deserialize, Serialize)]
// #[serde(rename(serialize = "o:idmap", deserialize = "idmap"))]
struct IdMap {
    #[serde(rename(serialize = "@v:ext", deserialize = "@ext"))]
    v_ext: String,
    #[serde(rename = "@data")]
    data: u32,
}