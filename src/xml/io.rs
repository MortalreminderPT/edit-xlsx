use std::io;
use std::path::Path;

pub(crate) trait Io<T> {
    fn from_path<P: AsRef<Path>>(file_path: P) -> io::Result<T>;
    fn save<P: AsRef<Path>>(&self, file_path: P);
}