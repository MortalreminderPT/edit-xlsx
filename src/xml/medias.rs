use std::path::{Path, PathBuf};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::manage::Io;

#[derive(Debug)]
pub(crate) struct Medias {
    medias: Vec<Media>
}

impl Medias {
    pub(crate) fn add_media<P: AsRef<Path>>(&mut self, file_path: P) -> u32 {
        let id = 1 + self.medias.len() as u32;
        let media = Media::new(id, &file_path);
        self.medias.push(media);
        id
    }
}

impl Io<Medias> for Medias {
    fn from_path<P: AsRef<Path>>(file_path: P) -> std::io::Result<Medias> {
        let mut medias = Medias { medias: vec![] };
        let mut id = 1;
        Ok(loop {
            let file = XlsxFileReader::from_path(&file_path, XlsxFileType::Medias(format!("image{id}.png")));
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

    fn save<P: AsRef<Path>>(&mut self, file_path: P) {
        self.medias.iter_mut().for_each(|m| { m.save(&file_path) });
    }
}

#[derive(Debug)]
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
    fn from_path<P: AsRef<Path>>(file_path: P) -> std::io::Result<Media> {
        // let media = Media::from_path(file_path);
        // Ok(media)
        todo!()
    }

    fn save<P: AsRef<Path>>(&mut self, file_path: P) {
        XlsxFileWriter::copy_from(&file_path, XlsxFileType::Medias(format!("image{}.png", self.id)), &self.file_path).unwrap();
    }
}