use std::fs;
use quick_xml::{de, se};
use crate::xml::drawings::{Drawings};

// todo: When encountering content that cannot be parsed, insert it to the file.

pub(crate) trait Insert {
    fn tag(&self) -> String;
    fn serialize_field(&self) -> String;
    fn insert(&self, content: &mut String) {
        let mut count = 1;
        let mut id = content.find(&self.tag()).unwrap_or_default();
        for c in content[id..content.len()].as_bytes() {
            id += 1;
            if *c == '<' as u8 {
                count += 1;
            } else if *c == '>' as u8 {
                count -= 1;
                println!("{c}");
                if count < 1 {
                    break;
                }
            }
        }
        content.insert_str(id, &self.serialize_field());
    }
}

impl Insert for Drawings {
    fn tag(&self) -> String {
        "xdr:wsDr".to_string()
    }

    fn serialize_field(&self) -> String {
        se::to_string_with_root("xdr:twoCellAnchor", &self.drawing).unwrap()
    }
}

// #[test]
// fn test() {
//     let xml = fs::read_to_string("tests/drawings2.txt").unwrap();
//     let drawings: Drawings = de::from_str(&xml).unwrap();
//     let mut text = fs::read_to_string("tests/drawings.txt").unwrap();
//     drawings.insert(&mut text);
//     fs::write("tests/drawings3.txt", text).unwrap();
//     // println!("{:?}", text);
// }