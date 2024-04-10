use std::cmp;
use std::fmt::Debug;
use serde::{Deserialize, Deserializer, Serialize, Serializer};
use crate::core::internal_tree::InternalTree;
use crate::result::ColResult;

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
pub(crate) struct Cols {
    #[serde(skip)]
    col: Vec<Col>,
    #[serde(rename = "col")]
    col_tree: InternalTree<Col>
}

impl Cols {
    pub(crate) fn new_col() -> Col {
        Col::default()
    }

    pub(crate) fn update_col_tree(&mut self, min: u32, max: u32, col: Col) {
        self.col_tree.update(min as i32, max as i32 + 1, &col).unwrap();
    }
}

impl Cols {
    pub(crate) fn update_col(&mut self, min: u32, max: u32, width: Option<f64>, style: Option<u32>, hidden: Option<u8>, best_fit: Option<u8>) -> ColResult<()> {
        let col = self.get_or_new_col(min, max);
        col.update_width(width);
        col.style = style;
        col.hidden = hidden;
        col.best_fit = best_fit;
        Ok(())
    }

    pub(crate) fn get_or_new_col(&mut self, col_min: u32, col_max: u32) -> &mut Col {
        let len = self.col.len();
        for i in 0..len {
            if self.col[i].min == col_min && self.col[i].max == col_max {
                return &mut self.col[i];
            }
        }
        for i in 0..len {
            if self.col[i].min <= col_min && self.col[i].max >= col_max {
                let mut col = Col::from(self.col[i]);
                col.max = col_max;
                col.min = col_min;
                self.col.push(col);
                return self.col.last_mut().unwrap();
            }
        }
        self.col.push(Col::new(col_min, col_max));
        self.col.last_mut().unwrap()
    }

    pub(crate) fn get_default_style(&self, col: u32) -> Option<u32> {
        let col = self.col.iter().filter(|c| c.min <= col && c.max >= col).last();
        if let Some(col) = col {
            return col.style;
        }
        None
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.col.is_empty()
    }
}

#[derive(Debug, Deserialize, Serialize, PartialEq, Default, Copy, Clone)]
pub(crate) struct Col {
    #[serde(rename = "@min")]
    min: u32,
    #[serde(rename = "@max")]
    max: u32,
    #[serde(rename = "@width", skip_serializing_if = "Option::is_none")]
    pub(crate) width: Option<f64>,
    #[serde(rename = "@style", skip_serializing_if = "Option::is_none")]
    pub(crate) style: Option<u32>,
    #[serde(rename = "@bestFit", skip_serializing_if = "Option::is_none")]
    pub(crate) best_fit: Option<u8>,
    #[serde(rename = "@hidden", skip_serializing_if = "Option::is_none")]
    pub(crate) hidden: Option<u8>,
    #[serde(rename = "@outlineLevel", skip_serializing_if = "Option::is_none")]
    pub(crate) outline_level: Option<u32>,
    #[serde(rename = "@collapsed", skip_serializing_if = "Option::is_none")]
    pub(crate) collapsed: Option<u8>,
    #[serde(rename = "@customWidth", skip_serializing_if = "Option::is_none")]
    pub(crate) custom_width: Option<u8>,
}

impl Col {
    fn new(min: u32, max: u32) -> Col {
        Col {
            min,
            max,
            width: None,
            style: None,
            best_fit: None,
            hidden: None,
            outline_level: None,
            collapsed: None,
            custom_width: None,
        }
    }

    fn update_width(&mut self, width: Option<f64>) {
        self.width = width;
        if let Some(_) = width {
            self.custom_width = Some(1);
        }
    }

    fn intersect(&self, col: &Col) -> Option<Col> {
        let min = cmp::max(self.min, col.min);
        let max = cmp::min(self.max, col.max);
        if min > max {
            return None;
        } else {
            let mut new_col = self.clone();
            new_col.min = min;
            new_col.max = max;
            new_col.fill_none(col);
            Some(new_col)
        }
    }

    fn fill_none(&mut self, col: &Col) {
        if let None = self.width {
            self.width = col.width;
        }
        if let None = self.style {
            self.style = col.style;
        }
        if let None = self.best_fit {
            self.best_fit = col.best_fit;
        }
        if let None = self.hidden {
            self.hidden = col.hidden;
        }
        if let None = self.outline_level {
            self.outline_level = col.outline_level;
        }
        if let None = self.collapsed {
            self.collapsed = col.collapsed;
        }
        if let None = self.custom_width {
            self.custom_width = col.custom_width;
        }
    }
}

impl Serialize for InternalTree<Col> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error> where S: Serializer {
        let mut v: Vec<(i32, i32, Col)> = self.to_vec();
        let cols: Vec<Col> = v.iter_mut()
            .map(|(l, r, c)| {
                c.min = *l as u32;
                c.max = *r as u32 - 1;
                *c
            })
            // .map(|v| v.2)
            .collect();
        serializer.serialize_some(&cols)
    }
}

impl<'de> Deserialize<'de> for InternalTree<Col> {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error> where D: Deserializer<'de> {
        // let v: Vec<T> = Vec::new();
        let col: Vec<Col> = Deserialize::deserialize(deserializer)?;
        let mut v: Vec<(i32, i32, Col)> = Vec::new();
        col.iter().for_each(|&c|v.push((c.min as i32, 1 + c.max as i32, c)));
        // let v: Vec<(i32, i32, Col> = v.iter().map(|&v|(v.min, v.max + 1, v)).collect();
        let tree = InternalTree::from_vec(&v);
        Ok(tree)
    }
}