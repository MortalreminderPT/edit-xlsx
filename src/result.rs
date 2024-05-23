use std::{error, fmt, io};
use quick_xml::DeError;
use zip::result::ZipError;

pub type CellResult<T> = Result<T, CellError>;
#[derive(Debug)]
pub enum CellError {
    CellNotFound
}

impl fmt::Display for CellError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match *self {
            CellError::CellNotFound => write!(f, "Cell not found"),
        }
    }
}

impl error::Error for CellError {
    fn source(&self) -> Option<&(dyn error::Error + 'static)> {
        match *self {
            CellError::CellNotFound => None,
        }
    }
}

pub type RowResult<T> = Result<T, RowError>;
#[derive(Debug)]
pub enum RowError {
    RowNotFound,
    CellError(CellError),
}

impl fmt::Display for RowError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match *self {
            RowError::RowNotFound => write!(f, "Row not found"),
            RowError::CellError(ref err) => write!(f, "Cell Error: {:?}", err),
        }
    }
}

impl error::Error for RowError {
    fn source(&self) -> Option<&(dyn error::Error + 'static)> {
        match *self {
            RowError::RowNotFound => None,
            RowError::CellError(ref err) => Some(err),
        }
    }
}


impl From<CellError> for RowError { fn from(err: CellError) -> RowError { RowError::CellError(err) } }

pub type ColResult<T> = Result<T, ColError>;
#[derive(Debug)]
pub enum ColError {
    ColNotFound,
}

impl fmt::Display for ColError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match *self {
            ColError::ColNotFound => write!(f, "Column not found"),
        }
    }
}

impl error::Error for ColError {
    fn source(&self) -> Option<&(dyn error::Error + 'static)> {
        match *self {
            ColError::ColNotFound => None,
        }
    }
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


impl fmt::Display for WorkSheetError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match *self {
            WorkSheetError::Io(ref err) => write!(f, "I/O Error: {:?}", err),
            WorkSheetError::DeError(ref err) => write!(f, "Deserialization Error: {:?}", err),
            WorkSheetError::ZipError(ref err) => write!(f, "ZIP File Error: {:?}", err),
            WorkSheetError::FileNotFound => write!(f, "File not found"),
            WorkSheetError::RowError(ref err) => write!(f, "Row Error: {:?}", err),
            WorkSheetError::ColError(ref err) => write!(f, "Column Error: {:?}", err),
            WorkSheetError::DuplicatedSheets => write!(f, "Duplicated Sheets"),
            WorkSheetError::FormatError => write!(f, "Format Error"),
        }
    }
}

impl error::Error for WorkSheetError {
    fn source(&self) -> Option<&(dyn error::Error + 'static)> {
        match *self {
            WorkSheetError::Io(ref err) => Some(err),
            WorkSheetError::DeError(ref err) => Some(err),
            WorkSheetError::ZipError(ref err) => Some(err),
            WorkSheetError::FileNotFound => None,
            WorkSheetError::RowError(ref err) => Some(err),
            WorkSheetError::ColError(ref err) => Some(err),
            WorkSheetError::DuplicatedSheets => None,
            WorkSheetError::FormatError => None,
        }
    }
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



impl fmt::Display for WorkbookError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match *self {
            WorkbookError::Io(ref err) => write!(f, "I/O Error: {:?}", err),
            WorkbookError::ZipError(ref err) => write!(f, "ZIP File Error: {:?}", err),
            WorkbookError::SheetError(ref err) => write!(f, "Worksheet Error: {:?}", err),
            WorkbookError::FileNotFound => write!(f, "File not found"),
            WorkbookError::RelationshipError(ref err) => write!(f, "Relationship Error: {:?}", err),
        }
    }
}

impl error::Error for WorkbookError {
    fn source(&self) -> Option<&(dyn error::Error + 'static)> {
        match *self {
            WorkbookError::Io(ref err) => Some(err),
            WorkbookError::ZipError(ref err) => Some(err),
            WorkbookError::SheetError(ref err) => Some(err),
            WorkbookError::FileNotFound => None,  // No underlying source error
            WorkbookError::RelationshipError(ref err) => Some(err),
        }
    }
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

impl fmt::Display for RelationshipError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match *self {
            RelationshipError::UnsupportedNamespace => write!(f, "Unsupported namespace"),
        }
    }
}

impl error::Error for RelationshipError {
    fn source(&self) -> Option<&(dyn error::Error + 'static)> {
        None
    }
}

