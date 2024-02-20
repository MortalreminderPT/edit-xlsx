use std::io;
use quick_xml::DeError;
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

pub type ColResult<T> = Result<T, ColError>;
#[derive(Debug)]
pub enum ColError {
    ColNotFound,
}

pub type SheetResult<T> = Result<T, SheetError>;
#[derive(Debug)]
pub enum SheetError {
    Io(io::Error),
    DeError(DeError),
    ZipError(ZipError),
    FileNotFound,
    RowError(RowError),
    ColError(ColError),
    DuplicatedSheets,
}

impl From<DeError> for SheetError { fn from(err: DeError) -> SheetError { SheetError::DeError(err) } }
impl From<io::Error> for SheetError { fn from(err: io::Error) -> SheetError {
        SheetError::Io(err)
    } }
impl From<RowError> for SheetError { fn from(err: RowError) -> SheetError { SheetError::RowError(err) } }
impl From<ColError> for SheetError { fn from(err: ColError) -> SheetError { SheetError::ColError(err) } }

pub type WorkbookResult<T> = Result<T, WorkbookError>;
#[derive(Debug)]
pub enum WorkbookError {
    Io(io::Error),
    ZipError(ZipError),
    SheetError(SheetError),
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

impl From<SheetError> for WorkbookError {
    fn from(err: SheetError) -> WorkbookError {
        WorkbookError::SheetError(err)
    }
}