use crate::utils::col_helper;

#[derive(Default, Debug, Copy, Clone)]
pub(crate) struct CellReference {
    row: u32,
    col: u32,
    abs_row: bool,
    abs_col: bool,
}

impl CellReference {
    fn new(row: u32, col: u32, abs_row: bool, abs_col: bool) -> CellReference {
        CellReference {
            row,
            col,
            abs_row,
            abs_col,
        }
    }
    fn from_str(refer: &str) -> CellReference {
        let mut coord = CellReference::default();
        let mut row = String::new();
        let mut col = String::new();
        let mut dollar = None;
        for c in refer.chars() {
            match c {
                '$' => { let _ = dollar.insert(1); },
                'A'..='Z' | 'a'..='z' => {
                    col.push(c);
                    if dollar.take().is_some() {
                        coord.abs_col = true;
                    }
                },
                '0'..='9' => {
                    row.push(c);
                    if dollar.take().is_some() {
                        coord.abs_row = true;
                    }
                },
                _ => {},
            }
        }
        coord.row = row.parse().unwrap_or_default();
        coord.col = col_helper::to_col(&col);
        coord
    }
    fn to_string(&self) -> String {
        let mut refer = String::new();
        if self.abs_col {
            refer.push('$')
        }
        refer.push_str(&col_helper::to_col_name(self.col));
        if self.abs_row {
            refer.push('$')
        }
        refer.push_str(&self.row.to_string());
        refer
    }
}