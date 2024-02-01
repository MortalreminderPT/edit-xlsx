pub type CellResult<T> = Result<T, CellError>;
#[derive(Debug)]
pub enum CellError {
    CellNotFound
}

pub type RowResult<T> = Result<T, RowError>;
#[derive(Debug)]
pub enum RowError {
    RowNotFound,
    CellError(CellError),
}
