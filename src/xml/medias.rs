use std::io;
use std::path::{Path, PathBuf};
use zip::read::ZipFile;
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::io::Io;

#[derive(Debug, Default)]
pub(crate) struct Medias {
    medias: Vec<Media>
}

impl Medias {
    pub(crate) fn add_media<P: AsRef<Path>>(&mut self, file_path: P) -> u32 {
        // let id = 1 + self.medias.len() as u32;
        let id = 1 + self.max_id(); // self.medias.iter().map(|m| m.id).max().unwrap();
        let media = Media::new(id, &file_path);
        self.medias.push(media);
        // self.medias.iter().max_by(|m|m.id);
        id
    }

    fn max_id(&self) -> u32 {
        self.medias.iter().map(|m| m.id).max().unwrap_or_default()
    }

    pub(crate) fn add_existed_media(&mut self, file_name: &str) -> u32 {
        let id: u32 = file_name
            .chars()
            .filter(|&i| i >= '0' && i <= '9')
            .collect::<String>()
            .parse()
            .unwrap_or(1 + self.max_id());
        self.medias.push(Media::by_id(id));
        id
    }
}

impl Io<Medias> for Medias {
    // fn from_path<P: AsRef<Path>>(file_path: P) -> io::Result<Medias> {
    //     let mut medias = Medias { medias: vec![] };
    //     let mut id = 1;
    //     Ok(loop {
    //         let extension = &file_path.as_ref().extension().unwrap_or("png".as_ref()).to_string_lossy();
    //         let file_name = format!("image{}.{}", id, extension);
    //         let file = XlsxFileReader::from_path(&file_path, XlsxFileType::Medias(file_name));
    //         match file {
    //             Ok(file) => {
    //                 let media = Media::new(id, file.file_path);
    //                 medias.medias.push(media);
    //                 id += 1;
    //             },
    //             Err(_) => {
    //                 break medias;
    //             },
    //         }
    //     })
    // }

    fn save<P: AsRef<Path>>(& self, file_path: P) {
        self.medias.iter().for_each(|m| { m.save(&file_path) });
    }
}

#[derive(Debug, Default)]
struct Media {
    id: u32,
    file_path: Option<PathBuf>,
}

impl Media {
    fn new<P: AsRef<Path>>(id: u32, file_path: P) -> Media {
        Media {
            id,
            file_path: Some(file_path.as_ref().to_path_buf()),
        }
    }

    fn by_id(id: u32) -> Media {
        Media {
            id,
            file_path: None,
        }
    }
}

impl Io<Media> for Media {
    fn save<P: AsRef<Path>>(&self, file_path: P) {
        if let Some(path) = &self.file_path {
            let extension = path.extension().unwrap_or("png".as_ref()).to_string_lossy();
            let file_name = format!("image{}.{}", self.id, extension);
            XlsxFileWriter::copy_from(&file_path, XlsxFileType::Medias(file_name), &path).unwrap();
        }
    }
}