use std::fs::File;
use std::{fs, io};
use std::io::{Read, Write};
use std::path::{Path, PathBuf};

pub enum XlsxFileType {
    WorkbookFile,
    SheetFile(u32),
    SharedStringFile,
    StylesFile,
    WorkbookRels,
    WorksheetRels(u32),
    ContentTypes,
    Medias(String),
}

pub struct XlsxFileReader {
    file_type: XlsxFileType,
    pub(crate) file_path: PathBuf,
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
        let file_path = parse_path(&base_path, &file_type);
        Ok(XlsxFileWriter {
            file: {
                Self::mkdir(&base_path, &file_type)?;
                File::create(&file_path)?
            },
            file_type,
            file_path,
        })
    }

    fn mkdir<P: AsRef<Path>>(base_path: P, file_type: &XlsxFileType) -> io::Result<()> {
        let file_path = parse_path(base_path, &file_type);
        let mut dirs = file_path.clone();
        dirs.pop();
        fs::create_dir_all(&dirs).unwrap_or_else(|_| {});
        Ok(())
    }

    pub(crate) fn write_all(&mut self, buf: &[u8]) -> io::Result<()> {
        self.file.write_all(buf)
    }

    pub(crate) fn copy_from<P, Q>(base_path: P, file_type: XlsxFileType, from: Q) -> io::Result<()>
        where P: AsRef<Path>,
              Q: AsRef<Path>
    {
        let file_path = parse_path(&base_path, &file_type);
        Self::mkdir(&base_path, &file_type)?;
        if from.as_ref().to_path_buf() == file_path.to_path_buf() {
            return Ok(())
        }
        fs::copy(from, file_path)?;
        Ok(())
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
        },
        XlsxFileType::WorksheetRels(id) => {
            base_path.as_ref().join(format!("xl/worksheets/_rels/sheet{id}.xml.rels"))
        },
        XlsxFileType::ContentTypes => {
            base_path.as_ref().join("[Content_Types].xml")
        }
        XlsxFileType::Medias(name) => {
            base_path.as_ref().join(format!("xl/media/{name}"))
        }
    }
}