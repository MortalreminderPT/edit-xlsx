use std::io;
use zip::result::ZipError;

pub type SheetResult<T> = Result<T, SheetError>;

#[derive(Debug)]
pub enum SheetError {
    Io(io::Error),
    ZipError(ZipError),
    FileNotFound,
    RowError,
    ColError,
}

