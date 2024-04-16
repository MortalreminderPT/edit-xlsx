use std::{fs, io};
use std::fs::File;
use std::io::{Read, Write};
use std::path::Path;
use walkdir::WalkDir;
use zip::CompressionMethod;
use zip::result::ZipError;
use zip::write::FileOptions;

pub(crate) fn extract_dir<P: AsRef<Path>>(file_path: P, target: &str) -> zip::result::ZipResult<String> {
    // read file from file path
    let file = File::open(&file_path)?;
    let mut archive = zip::ZipArchive::new(file)?;
    // construct a base path for extracted files
    let base_path = Path::new(&target);
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
            fs::create_dir_all(&out_path)?;
        } else {
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
    let tmp_dir = base_path.to_str().ok_or(ZipError::FileNotFound)?.to_string();
    Ok(tmp_dir)
}

pub(crate) fn zip_dir<P: AsRef<Path>>(prefix: &str, file_path: P) -> zip::result::ZipResult<()> {
    let writer = File::create(&file_path)?;
    let walk_dir = WalkDir::new(&prefix);
    let it = walk_dir.into_iter();
    let it = &mut it.filter_map(|e| e.ok());

    let mut zip = zip::ZipWriter::new(writer);
    let options = FileOptions::default()
        .compression_method(CompressionMethod::Deflated)
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

#[test]
fn test() -> io::Result<()> {
    let file = File::open("./examples/xlsx/accounting.xlsx")?;
    // 创建 ZipArchive 对象
    let mut archive = zip::ZipArchive::new(file)?;
    let file_path = "xl/styles.xml";

    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        println!("{}", file.name());
        if file.name() == file_path {
            let mut contents = String::new();
            file.read_to_string(&mut contents)?;
            println!("File contents: {}", contents);
        }
    }
    // if let Some(mut file) = find_file(&mut archive, file_path)? {
    //     // 读取文件内容
    //     let mut contents = String::new();
    //     file.read_to_string(&mut contents)?;
    //     println!("File contents: {}", contents);
    // } else {
    //     println!("File not found: {}", file_path);
    // }
    Ok(())
}