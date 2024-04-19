use std::io::{Error, ErrorKind};

pub fn error_text(text: &str) -> Error {
    Error::new(ErrorKind::Other, text)
}
