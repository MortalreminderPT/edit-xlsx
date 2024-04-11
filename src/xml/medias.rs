use std::io;
use std::path::{Path, PathBuf};
use serde::Deserialize;
use zip::read::ZipFile;
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::io::Io;

#[derive(Debug, Default)]
pub(crate) struct Medias {
    medias: Vec<Media>
}

unsafe impl Send for Medias {}

impl Medias {
    pub(crate) fn add_media<P: AsRef<Path>>(&mut self, file_path: P) -> u32 {
        let id = 1 + self.medias.len() as u32;
        let media = Media::new(id, &file_path);
        self.medias.push(media);
        id
    }
}

impl Io<Medias> for Medias {
    fn from_path<P: AsRef<Path>>(file_path: P) -> io::Result<Medias> {
        let mut medias = Medias { medias: vec![] };
        let mut id = 1;
        Ok(loop {
            let extension = &file_path.as_ref().extension().unwrap_or("png".as_ref()).to_string_lossy();
            let file_name = format!("image{}.{}", id, extension);
            let file = XlsxFileReader::from_path(&file_path, XlsxFileType::Medias(file_name));
            match file {
                Ok(file) => {
                    let media = Media::new(id, file.file_path);
                    medias.medias.push(media);
                    id += 1;
                },
                Err(_) => {
                    break medias;
                },
            }
        })
    }

    fn save<P: AsRef<Path>>(& self, file_path: P) {
        self.medias.iter().for_each(|m| { m.save(&file_path) });
    }

    fn from_zip_file(file: &mut ZipFile) -> Medias {
        todo!()
    }
}

#[derive(Debug, Default)]
struct Media {
    id: u32,
    file_path: PathBuf,
}

impl Media {
    fn new<P: AsRef<Path>>(id: u32, file_path: P) -> Media {
        Media {
            id,
            file_path: file_path.as_ref().to_path_buf(),
        }
    }
}

impl Io<Media> for Media {
    fn from_path<P: AsRef<Path>>(_file_path: P) -> io::Result<Media> {
        todo!()
    }

    fn save<P: AsRef<Path>>(&self, file_path: P) {
        let extension = &self.file_path.extension().unwrap_or("png".as_ref()).to_string_lossy();
        let file_name = format!("image{}.{}", self.id, extension);
        XlsxFileWriter::copy_from(&file_path, XlsxFileType::Medias(file_name), &self.file_path).unwrap();
    }

    fn from_zip_file(file: &mut ZipFile) -> Media {
        todo!()
    }
}