pub(crate) mod vml_drawing;

use std::io;
use std::io::Read;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use zip::read::ZipFile;
use crate::api::cell::location::{Location, LocationRange};
use crate::api::relationship::Rel;
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename(serialize = "xdr:wsDr", deserialize = "wsDr"))]
pub(crate) struct Drawings {
    #[serde(rename(serialize = "@xmlns:xdr", deserialize = "@xmlns:xdr"), default, skip_serializing_if = "String::is_empty")]
    xmlns_xdr: String,
    #[serde(rename(serialize = "@xmlns:a", deserialize = "@xmlns:a"), default, skip_serializing_if = "String::is_empty")]
    xmlns_a: String,
    #[serde(rename(serialize = "xdr:twoCellAnchor", deserialize = "twoCellAnchor"), default)]
    drawing: Vec<Drawing>
}

impl Default for Drawings {
    fn default() -> Self {
        Self {
            xmlns_xdr: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing".to_string(),
            xmlns_a: "http://schemas.openxmlformats.org/drawingml/2006/main".to_string(),
            drawing: vec![],
        }
    }
}

impl Drawings {
    pub(crate) fn next_id(&self) -> u32 {
        let id = 1 + self.drawing.len() as u32;
        id
    }

    pub(crate) fn add_drawing<L: LocationRange>(&mut self, from_to: L, r_id: u32) {
        self.drawing.push(Drawing::new(from_to, r_id));
    }
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct Drawing {
    #[serde(rename = "@editAs")]
    edit_as: String,
    #[serde(rename(serialize = "xdr:from", deserialize = "from"))]
    from: DrawingLocation,
    #[serde(rename(serialize = "xdr:to", deserialize = "to"))]
    to: DrawingLocation,
    #[serde(rename(serialize = "xdr:pic", deserialize = "pic"))]
    pic: Picture,
    #[serde(rename(serialize = "xdr:clientData", deserialize = "clientData"))]
    client_data: ClientData,
}

impl Drawing {
    fn new<L: LocationRange>(from_to: L, r_id: u32) -> Drawing {
        let (from_row, from_col, to_row, to_col) = from_to.to_range();
        Drawing {
            edit_as: String::from("oneCell"),
            from: DrawingLocation::from_location((from_row, from_col)),
            to: DrawingLocation::from_location((to_row, to_col)),
            pic: Picture::from_id(r_id),
            client_data: ClientData::default(),
        }
    }
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct DrawingLocation {
    #[serde(rename(serialize = "xdr:col", deserialize = "col"))]
    col: u32,
    #[serde(rename(serialize = "xdr:colOff", deserialize = "colOff"))]
    col_off: u32,
    #[serde(rename(serialize = "xdr:row", deserialize = "row"))]
    row: u32,
    #[serde(rename(serialize = "xdr:rowOff", deserialize = "rowOff"))]
    row_off: u32
}

impl DrawingLocation {
    fn from_location<L: Location>(loc: L) -> DrawingLocation {
        let (row, col) = loc.to_location();
        DrawingLocation {
            col: col - 1,
            row: row - 1,
            col_off: 0,
            row_off: 0,
        }
    }
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct Picture {
    #[serde(rename(serialize = "xdr:nvPicPr", deserialize = "nvPicPr"))]
    pic_pr: PicPr,
    #[serde(rename(serialize = "xdr:blipFill", deserialize = "blipFill"))]
    blip_fill: BlipFill,
    #[serde(rename(serialize = "xdr:spPr", deserialize = "spPr"))]
    sp_pr: SpPr,
}

impl Picture {
    fn from_id(r_id: u32) -> Self {
        Self {
            pic_pr: PicPr::from_id(r_id),
            blip_fill: BlipFill::from_id(r_id),
            sp_pr: SpPr::default(),
        }
    }
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct PicPr {
    #[serde(rename(serialize = "xdr:cNvPr", deserialize = "cNvPr"))]
    c_nv_pr: CNvPr,
    #[serde(rename(serialize = "xdr:cNvPicPr", deserialize = "cNvPicPr"))]
    c_nv_pic_pr: CNvPicPr,
}

impl PicPr {
    fn from_id(id: u32) -> Self {
        Self {
            c_nv_pr: CNvPr::from_id(id),
            c_nv_pic_pr: Default::default(),
        }
    }
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct CNvPr {
    #[serde(rename = "@id")]
    id: String,
    #[serde(rename = "@name")]
    name: String,
}

impl CNvPr {
    fn from_id(id: u32) -> Self {
        Self {
            id: id.to_string(),
            name: format!("Picture {id}"),
        }
    }
}

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
struct CNvPicPr {
    #[serde(rename(serialize = "a:picLocks", deserialize = "picLocks"))]
    pic_locks: PicLocks,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct PicLocks {
    #[serde(rename = "@noChangeAspect")]
    no_change_aspect: u8,
}

impl Default for PicLocks {
    fn default() -> Self {
        Self {
            no_change_aspect: 1,
        }
    }
}

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
struct BlipFill {
    #[serde(rename(serialize = "a:blip", deserialize = "blip"))]
    blip: Blip,
    #[serde(rename(serialize = "a:stretch", deserialize = "stretch"))]
    stretch: Stretch,
}

impl BlipFill {
    fn from_id(r_id: u32) -> Self {
        Self {
            blip: Blip::from_id(r_id),
            stretch: Default::default(),
        }
    }
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct Blip {
    #[serde(rename(serialize = "@xmlns:r", deserialize = "@xmlns:r"), default, skip_serializing_if = "String::is_empty")]
    xmlns_r: String,
    #[serde(rename(serialize = "@r:embed", deserialize = "@embed"))]
    r_embed: Rel,
}

impl Default for Blip {
    fn default() -> Self {
        Self {
            xmlns_r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships".to_string(),
            r_embed: Default::default(),
        }
    }
}

impl Blip {
    fn from_id(r_id: u32) -> Blip {
        Blip {
            xmlns_r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships".to_string(),
            r_embed: Rel::from_id(r_id),
        }
    }
}

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
struct Stretch {
    #[serde(rename(serialize = "a:fillRect", deserialize = "fillRect"))]
    fill_rect: FillRect,
}

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
struct FillRect {}

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
struct SpPr {
    #[serde(rename(serialize = "a:prstGeom", deserialize = "prstGeom"))]
    prst_geom: PrstGeom,
}

#[derive(Debug, Clone ,Deserialize, Serialize)]
struct PrstGeom {
    #[serde(rename(serialize = "a:avLst", deserialize = "avLst"))]
    av_lst: AvLst,
    #[serde(rename(serialize = "@prst", deserialize = "@prst"))]
    prst: String,
}

impl Default for PrstGeom {
    fn default() -> Self {
        Self {
            av_lst: Default::default(),
            prst: "rect".to_string(),
        }
    }
}

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
struct AvLst {}

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
struct ClientData {}

impl Drawings {
    pub(crate) fn from_zip_file(mut file: &mut ZipFile) -> Self {
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        de::from_str(&xml).unwrap_or_default()
    }
}

impl Drawings {
    pub(crate) fn from_path<P: AsRef<Path>>(file_path: P, drawing_id: u32) -> io::Result<Drawings> {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::Drawings(drawing_id))?;
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let drawings: Drawings = de::from_str(&xml).unwrap();
        Ok(drawings)
    }

    pub(crate) fn save<P: AsRef<Path>>(& self, file_path: P, drawing_id: u32) {
        let xml = se::to_string_with_root("xdr:wsDr", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::Drawings(drawing_id)).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}