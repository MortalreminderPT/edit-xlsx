use std::io;
use zip::result::ZipError;
use crate::result::RowError;

pub type SheetResult<T> = Result<T, SheetError>;

#[derive(Debug)]
pub enum SheetError {
    Io(io::Error),
    ZipError(ZipError),
    FileNotFound,
    RowError(RowError),
    ColError,
}

impl From<RowError> for SheetError {
    fn from(err: RowError) -> SheetError {
        SheetError::RowError(err)
    }
}


