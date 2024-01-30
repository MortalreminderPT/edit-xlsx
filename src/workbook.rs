mod tests;
mod result;

use std::{fs, io};
use std::cell::RefCell;
use std::fs::File;
use std::io::{Read, Seek, Write};
use std::ops::Deref;
use std::path::Path;
use std::rc::{Rc, Weak};
use serde::{Deserialize, Serialize};
use walkdir::{DirEntry, WalkDir};
use zip::CompressionMethod;
use zip::write::FileOptions;
use crate::sheet::Sheet;
use crate::workbook::result::{WorkbookError, WorkbookResult};
use crate::shared_string::SharedString;

#[derive(Debug)]
pub struct Workbook {
    pub sheets: Vec<Sheet>,
    pub shared_string: Rc<RefCell<SharedString>>,
    pub tmp_path: String,
    pub file_path: String,
}

impl Workbook {
    pub fn get_mut_sheet(&mut self, id: u8) -> Option<&mut Sheet> {
        self.sheets.get_mut(usize::from(id))
    }

    pub fn get_sheet(&mut self, id: u8) -> Option<&Sheet> {
        self.sheets.get(usize::from(id))
    }
}

impl Workbook {
    pub fn from_path<P: AsRef<Path>>(file_path: P) -> Workbook {
        let tmp_path = Workbook::create_tmp_dir(&file_path).unwrap();
        let shared_string = SharedString::from_path(&tmp_path);
        let shared_string = Rc::new(RefCell::new(shared_string));
        let workbook = Workbook {
            sheets: vec![Sheet::from_path(&tmp_path, 1, Rc::clone(&shared_string))],
            shared_string: Rc::clone(&shared_string),
            tmp_path,
            file_path: file_path.as_ref().to_str().unwrap().to_string()
        };
        // add a ptr for sheets to workbook
        workbook
    }

    fn create_tmp_dir<P: AsRef<Path>>(file_path: P) -> WorkbookResult<String> {
        // parse the file name
        let file_name = file_path.as_ref().file_name().ok_or(WorkbookError::FileNotFound)?;
        // read file from file path
        let file = File::open(&file_path)?;
        let mut archive = zip::ZipArchive::new(file)?;
        // construct a base path for extracted files
        let binding = "./.editing-".to_owned() + file_name.to_str().unwrap();
        let base_path = Path::new(&binding);
        {
            match fs::create_dir(&base_path) {
                Err(why) => println!("! {:?}", why.kind()),
                Ok(_) => {},
            }
            for i in 0..archive.len() {
                let mut file = archive.by_index(i)?;
                let out_path = match file.enclosed_name() {
                    Some(path) => path.to_owned(),
                    None => continue,
                };
                let out_path = &base_path.join(out_path);
                {
                    let comment = file.comment();
                    if !comment.is_empty() {
                        println!("File {i} comment: {comment}");
                    }
                }
                if (*file.name()).ends_with('/') {
                    // println!("File {} extracted to \"{}\"", i, out_path.display());
                    fs::create_dir_all(&out_path)?;
                } else {
                    // println!(
                    //     "File {} extracted to \"{}\" ({} bytes)",
                    //     i,
                    //     out_path.display(),
                    //     file.size()
                    // );
                    if let Some(p) = out_path.parent() {
                        if !p.exists() {
                            fs::create_dir_all(p)?;
                        }
                    }
                    let mut outfile = fs::File::create(&out_path)?;
                    io::copy(&mut file, &mut outfile)?;
                }
                // Get and Set permissions
                #[cfg(unix)]
                {
                    use std::os::unix::fs::PermissionsExt;

                    if let Some(mode) = file.unix_mode() {
                        fs::set_permissions(&out_path, fs::Permissions::from_mode(mode))?;
                    }
                }
            }
        }
        Ok(base_path.to_str().unwrap().to_string())
    }
    pub fn save<P: AsRef<Path>>(&self, file_path: P) -> WorkbookResult<bool> {
        // save files
        self.shared_string.borrow_mut().save(&self.tmp_path);
        self.sheets[0].save(&self.tmp_path);

        // package files
        let tmp_path = &self.tmp_path;
        let file = File::create(&file_path).unwrap();
        let walk_dir = WalkDir::new(tmp_path);
        let it = walk_dir.into_iter();
        zip_dir(&mut it.filter_map(|e| e.ok()), tmp_path, file, CompressionMethod::Deflated)?;
        Ok(true)
    }
}

impl Drop for Workbook {
    fn drop(&mut self) {
        let droped = fs::remove_dir_all(&self.tmp_path);
    }
}

fn zip_dir<T>(
    it: &mut dyn Iterator<Item = DirEntry>,
    prefix: &str,
    writer: T,
    method: zip::CompressionMethod,
) -> zip::result::ZipResult<()>
    where
        T: Write + Seek,
{
    let mut zip = zip::ZipWriter::new(writer);
    let options = FileOptions::default()
        .compression_method(method)
        .unix_permissions(0o755);

    let mut buffer = Vec::new();
    for entry in it {
        let path = entry.path();
        let name = path.strip_prefix(Path::new(prefix)).unwrap();

        // Write file or directory explicitly
        // Some unzip tools unzip files with directory paths correctly, some do not!
        if path.is_file() {
            // println!("adding file {path:?} as {name:?} ...");
            #[allow(deprecated)]
            zip.start_file_from_path(name, options)?;
            let mut f = File::open(path)?;

            f.read_to_end(&mut buffer)?;
            zip.write_all(&buffer)?;
            buffer.clear();
        } else if !name.as_os_str().is_empty() {
            // Only if not root! Avoids path spec / warning
            // and mapname conversion failed error on unzip
            // println!("adding dir {path:?} as {name:?} ...");
            #[allow(deprecated)]
            zip.add_directory_from_path(name, options)?;
        }
    }
    zip.finish()?;
    Ok(())
}
