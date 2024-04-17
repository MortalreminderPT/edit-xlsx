use crate::utils::col_helper;

#[derive(Default, Debug, Copy, Clone)]
pub(crate) struct CellAddress {
    row: u32,
    col: u32,
}

impl CellAddress {
    fn new(row: u32, col: u32) -> CellAddress {
        CellAddress {
            row,
            col,
        }
    }
    fn from_str(refer: &str) -> CellAddress {
        let mut coord = CellAddress::default();
        let mut row = String::new();
        let mut col = String::new();
        for c in refer.chars() {
            match c {
                'A'..='Z' | 'a'..='z' => {
                    col.push(c);
                },
                '0'..='9' => {
                    row.push(c);
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
        refer.push_str(&col_helper::to_col_name(self.col));
        refer.push_str(&self.row.to_string());
        refer
    }
}