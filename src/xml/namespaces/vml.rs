use serde::{Deserialize, Serialize, Serializer};
use crate::xml::namespaces::excel as x;

///
/// xmlns:v="urn:schemas-microsoft-com:vml"
///

#[derive(Debug, Default, Deserialize, Serialize)]
#[serde(rename(serialize = "v:shapetype", deserialize = "shapetype"))]
pub(crate) struct ShapeType {
    #[serde(rename = "@id")]
    id: String,
    #[serde(rename = "@coordsize")]
    coord_size: String,
    #[serde(rename(serialize = "@o:spt", deserialize = "@spt"))]
    o_spt: String,
    #[serde(rename = "@path")]
    param_path: String,
    #[serde(rename(serialize = "v:stroke", deserialize = "stroke"))]
    stroke: Stroke,
    #[serde(rename(serialize = "v:path", deserialize = "path"))]
    path: Path,
}

#[derive(Debug, Default, Deserialize, Serialize)]
pub(crate) struct Shape {
    #[serde(rename = "@id")]
    id: String,
    #[serde(rename = "@type")]
    param_type: String,
    #[serde(rename = "@style")]
    style: String,
    #[serde(rename = "@fillcolor")]
    fillcolor: String,
    #[serde(rename(serialize = "@o:insetmode", deserialize = "@insetmode"))]
    o_insetmode: String,
    #[serde(rename(serialize = "v:fill", deserialize = "fill"))]
    fill: Fill,
    #[serde(rename(serialize = "v:shadow", deserialize = "shadow"))]
    shadow: Shadow,
    #[serde(rename(serialize = "v:path", deserialize = "path"))]
    path: Path,
    #[serde(rename(serialize = "v:textbox", deserialize = "textbox"))]
    textbox: TextBox,
    #[serde(rename(serialize = "x:ClientData", deserialize = "ClientData"))]
    clientdata: x::ClientData,
}

#[derive(Debug, Default, Deserialize, Serialize)]
struct Stroke {
    #[serde(rename = "@joinstyle")]
    join_style :String
}

#[derive(Debug, Default, Deserialize, Serialize)]
struct Path {
    #[serde(rename = "@gradientshapeok", skip_serializing_if = "Option::is_none")]
    gradient_shape_ok: Option<String>,
    #[serde(rename(serialize = "@o:connecttype", deserialize = "@connecttype"))]
    o_connect_type: String,
}

#[derive(Debug, Default, Deserialize, Serialize)]
#[serde(rename(serialize = "v:fill", deserialize = "fill"))]
struct Fill {
    #[serde(rename = "@color2")]
    color2: String,
}

#[derive(Debug, Default, Deserialize, Serialize)]
#[serde(rename(serialize = "v:shadow", deserialize = "shadow"))]
struct Shadow {
    #[serde(rename = "@on")]
    on: String,
    #[serde(rename = "@color")]
    color: String,
    #[serde(rename = "@obscured")]
    obscured: String,
}

#[derive(Debug, Default, Deserialize, Serialize)]
#[serde(rename(serialize = "v:textbox", deserialize = "textbox"))]
struct TextBox {
    #[serde(rename = "@style")]
    style: String,
    div: Div
}

#[derive(Debug, Default, Deserialize, Serialize)]
#[serde(rename = "div")]
struct Div {
    #[serde(rename = "@style")]
    style: String,
}