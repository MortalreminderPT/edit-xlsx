use std::fs::File;
use std::{fs, io};
use std::io::{Read, Write};
use std::path::{Path, PathBuf};

pub enum XlsxFileType {
    WorkbookFile,
    SheetFile(String),
    SharedStringFile,
    StylesFile,
    WorkbookRels,
    WorksheetRels(u32),
    ContentTypes,
    Medias(String),
    Drawings(u32),
    DrawingRels(u32),
    MetaData,
    CoreProperties,
    AppProperties,
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
        // let file_path = get_path(base_path, &file_type);
        let file_path = file_type.get_path(base_path);
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
        let file_path = file_type.get_path(&base_path);
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
        let file_path = file_type.get_path(base_path);
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
        let file_path = file_type.get_path(&base_path);
        Self::mkdir(&base_path, &file_type)?;
        if from.as_ref().to_path_buf() == file_path.to_path_buf() {
            return Ok(())
        }
        fs::copy(from, file_path)?;
        Ok(())
    }
}

impl XlsxFileType {
    fn get_dir(&self) -> &str {
        match self {
            XlsxFileType::WorkbookFile | XlsxFileType::SharedStringFile | XlsxFileType::StylesFile | XlsxFileType::MetaData => "./xl",
            // XlsxFileType::SheetFile(_) => "./xl/worksheets",
            XlsxFileType::SheetFile(_) => "./xl",
            XlsxFileType::WorkbookRels => "./xl/_rels",
            XlsxFileType::WorksheetRels(_) => "./xl/worksheets/_rels",
            XlsxFileType::ContentTypes => ".",
            XlsxFileType::Medias(_) => "./xl/media",
            XlsxFileType::Drawings(_) => "./xl/drawings",
            XlsxFileType::DrawingRels(_) => "./xl/drawings/_rels",
            XlsxFileType::CoreProperties | XlsxFileType::AppProperties => "./docProps",
        }
    }
    fn get_filename(&self) -> String {
        match self {
            XlsxFileType::WorkbookFile => "workbook.xml".to_string(),
            // XlsxFileType::SheetFile(id) => format!("sheet{id}.xml"),
            XlsxFileType::SheetFile(target) => format!("{target}"), //format!("sheet{id}.xml"),
            XlsxFileType::SharedStringFile => "sharedStrings.xml".to_string(),
            XlsxFileType::StylesFile => "styles.xml".to_string(),
            XlsxFileType::WorkbookRels => "workbook.xml.rels".to_string(),
            XlsxFileType::WorksheetRels(id) => format!("sheet{id}.xml.rels"),
            XlsxFileType::ContentTypes => "[Content_Types].xml".to_string(),
            XlsxFileType::Medias(name) => format!("{name}"),
            XlsxFileType::Drawings(id) => format!("drawing{id}.xml"),
            XlsxFileType::DrawingRels(id) => format!("drawing{id}.xml.rels"),
            XlsxFileType::MetaData => "metadata.xml".to_string(),
            XlsxFileType::CoreProperties => "core.xml".to_string(),
            XlsxFileType::AppProperties => "app.xml".to_string(),
        }
    }
    pub(crate) fn get_path<P: AsRef<Path>>(&self, base_path: P) -> PathBuf {
        base_path.as_ref().join(self.get_dir()).join(self.get_filename())
    }
}