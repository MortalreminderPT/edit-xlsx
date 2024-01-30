use std::io;
use zip::result::ZipError;

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