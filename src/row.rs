use serde::{Deserialize, Serialize};
use crate::cell::Cell;

pub struct Row {
    cells: Vec<Cell>
}