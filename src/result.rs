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

impl From<CellError> for RowError { fn from(err: CellError) -> RowError { RowError::CellError(err) } }

pub type ColResult<T> = Result<T, ColError>;
#[derive(Debug)]
pub enum ColError {
    ColNotFound,
}

pub type WorkSheetResult<T> = Result<T, WorkSheetError>;
#[derive(Debug)]
pub enum WorkSheetError {
    Io(io::Error),
    DeError(DeError),
    ZipError(ZipError),
    FileNotFound,
    RowError(RowError),
    ColError(ColError),
    DuplicatedSheets,
    FormatError,
}

impl From<DeError> for WorkSheetError { fn from(err: DeError) -> WorkSheetError { WorkSheetError::DeError(err) } }
impl From<io::Error> for WorkSheetError { fn from(err: io::Error) -> WorkSheetError {
        WorkSheetError::Io(err)
    } }
impl From<RowError> for WorkSheetError { fn from(err: RowError) -> WorkSheetError { WorkSheetError::RowError(err) } }
impl From<ColError> for WorkSheetError { fn from(err: ColError) -> WorkSheetError { WorkSheetError::ColError(err) } }

pub type WorkbookResult<T> = Result<T, WorkbookError>;
#[derive(Debug)]
pub enum WorkbookError {
    Io(io::Error),
    ZipError(ZipError),
    SheetError(WorkSheetError),
    FileNotFound,
    RelationshipError(RelationshipError),
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

impl From<WorkSheetError> for WorkbookError {
    fn from(err: WorkSheetError) -> WorkbookError {
        WorkbookError::SheetError(err)
    }
}

pub type RelationshipResult<T> = Result<T, RelationshipError>;

#[derive(Debug)]
pub enum RelationshipError {
    UnsupportedNamespace,
}

