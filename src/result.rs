use std::io;
use zip::result::ZipError;

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


pub type WorkbookResult<T> = Result<T, WorkbookError>;
#[derive(Debug)]
pub enum WorkbookError {
    Io(io::Error),
    ZipError(ZipError),
    FileNotFound,
}

impl From<io::Error> for WorkbookError {
    fn from(err: io::Error) -> WorkbookError {
        WorkbookError::Io(err)
    }
}

impl From<ZipError> for WorkbookError {
    fn from(err: ZipError) -> WorkbookError {
        WorkbookError::ZipError(err)
    }
}