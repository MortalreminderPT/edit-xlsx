use serde::{Deserialize, Serialize};

///
/// xmlns:x="urn:schemas-microsoft-com:office:excel"
///

#[derive(Debug, Deserialize, Serialize)]
#[serde(rename(serialize = "x:ClientData", deserialize = "ClientData"))]
pub(crate) struct ClientData {
    #[serde(rename(serialize = "@ObjectType", deserialize = "@ObjectType"))]
    object_type: String,
    #[serde(rename(serialize = "x:MoveWithCells", deserialize = "MoveWithCells"))]
    move_with_cells: MoveWithCells,
    #[serde(rename(serialize = "x:SizeWithCells", deserialize = "SizeWithCells"))]
    size_with_cells: SizeWithCells,
    #[serde(rename(serialize = "x:Anchor", deserialize = "Anchor"))]
    anchor: Anchor,
    #[serde(rename(serialize = "x:Row", deserialize = "Row"))]
    row: Row,
    #[serde(rename(serialize = "x:Column", deserialize = "Column"))]
    column: Column,
}

#[derive(Debug, Deserialize, Serialize)]
struct MoveWithCells {}

#[derive(Debug, Deserialize, Serialize)]
struct SizeWithCells {
}

#[derive(Debug, Deserialize, Serialize)]
struct Anchor {
    #[serde(rename = "$value")]
    value: String,// Vec<u32>,
}
#[derive(Debug, Deserialize, Serialize)]
#[serde(rename(serialize = "x:AutoFill", deserialize = "AutoFill"))]
struct AutoFill {
    #[serde(rename = "$value")]
    value: bool,
}
#[derive(Debug, Deserialize, Serialize)]
struct Row {
    #[serde(rename = "$value")]
    value: u32,
}
#[derive(Debug, Deserialize, Serialize)]
struct Column {
    #[serde(rename = "$value")]
    value: u32,
}
