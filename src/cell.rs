

pub struct Cell {
    col_id: u32,
    row_id: u32,
    value: CellValue,
    style: CellStyle,
}

enum CellValue {
    Blank,
    Numeric(f64),
    Text(String),
    DateTime,
    Formula,
    Boolean(bool),
    Err,
}

struct CellStyle {
    
}