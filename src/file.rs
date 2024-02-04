use std::fs::File;
use std::io;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};

pub enum XlsxFileType {
    WorkbookFile,
    SheetFile(u32),
    SharedStringFile,
    StylesFile,
    WorkbookRels,
}

pub struct XlsxFileReader {
    file_type: XlsxFileType,
    file_path: PathBuf,
    file: File,
}

pub struct XlsxFileWriter {
    file_type: XlsxFileType,
    file_path: PathBuf,
    file: File,
}

impl XlsxFileReader {
    pub(crate) fn from_path<P: AsRef<Path>>(base_path: P, file_type: XlsxFileType) -> io::Result<XlsxFileReader> {
        let file_path = parse_path(base_path, &file_type);
        Ok(XlsxFileReader {
            file: File::open(&file_path)?,
            file_type,
            file_path,
        })
    }

    pub(crate) fn read_to_string(&mut self, target_string: &mut String) -> io::Result<usize> {
        self.file.read_to_string(target_string)
    }
}

impl XlsxFileWriter {
    pub(crate) fn from_path<P: AsRef<Path>>(base_path: P, file_type: XlsxFileType) -> io::Result<XlsxFileWriter> {
        let file_path = parse_path(base_path, &file_type);
        Ok(XlsxFileWriter {
            file: File::create(&file_path)?,
            file_type,
            file_path,
        })
    }

    pub(crate) fn write_all(&mut self, mut buf: &[u8]) -> io::Result<()> {
        self.file.write_all(buf)
    }
}

fn parse_path<P: AsRef<Path>>(base_path: P, file_type: &XlsxFileType) -> PathBuf {
    match file_type {
        XlsxFileType::WorkbookFile => {
            base_path.as_ref().join("xl/workbook.xml")
        }
        XlsxFileType::SheetFile(id) => {
            base_path.as_ref().join(format!("xl/worksheets/sheet{id}.xml"))
        }
        XlsxFileType::SharedStringFile => {
            base_path.as_ref().join("xl/sharedStrings.xml")
        }
        XlsxFileType::StylesFile => {
            base_path.as_ref().join("xl/styles.xml")
        }
        XlsxFileType::WorkbookRels => {
            base_path.as_ref().join("xl/_rels/workbook.xml.rels")
        }
    }
}