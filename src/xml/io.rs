use std::io;
use std::io::Read;
use std::path::Path;
use quick_xml::de;
use zip::read::ZipFile;

pub(crate) trait Io<T: Default> {
    fn from_path<P: AsRef<Path>>(file_path: P) -> io::Result<T>;
    fn save<P: AsRef<Path>>(&self, file_path: P);
    async fn from_path_async<P: AsRef<Path>>(file_path: P) -> io::Result<T> {
        Self::from_path(file_path)
    }
    async fn save_async<P: AsRef<Path>>(&self, file_path: P) {
        self.save(file_path)
    }

    fn from_zip_file(file: &mut ZipFile) -> T;
}