pub(crate) mod font;
pub(crate) mod border;
pub(crate) mod fill;

use std::collections::{HashMap, HashSet};
use std::hash::Hash;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::Format;
use crate::xml::common::{ExtLst, XmlnsAttrs};
use crate::xml::manage::XmlIo;
use crate::xml::style::border::Border;
use crate::xml::style::fill::Fill;
use crate::xml::style::font::Font;

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct StyleSheet {
    #[serde(flatten)]
    xmlns_attrs: XmlnsAttrs,
    #[serde(rename = "fonts")]
    fonts: Fonts,
    #[serde(rename = "fills")]
    fills: Fills,
    #[serde(rename = "borders")]
    borders: Borders,
    #[serde(rename = "cellStyleXfs")]
    cell_style_xfs: CellStyleXfs,
    #[serde(rename = "cellXfs")]
    cell_xfs: CellXfs,
    #[serde(rename = "cellStyles")]
    cell_styles: CellStyles,
    #[serde(rename = "dxfs")]
    dxfs: Dxfs,
    #[serde(rename = "tableStyles")]
    table_styles: TableStyles,
    #[serde(rename = "extLst", skip_serializing_if = "Option::is_none")]
    ext_lst: Option<ExtLst>,
}

#[derive(Debug, Deserialize, Serialize)]
struct Fonts {
    #[serde(rename = "@count", default)]
    count: u32,
    #[serde(rename(serialize = "@x14ac:knownFonts", deserialize = "@knownFonts"), default)]
    x14ac_known_fonts: u32,
    #[serde(rename = "font", default)]
    fonts: Vec<Font>,
}

impl Fonts {
    pub(crate) fn add_font(&mut self, font: &Font) -> u32 {
        for i in 0..self.fonts.len() {
            if self.fonts[i] == *font {
                return i as u32;
            }
        }
        self.count += 1;
        self.fonts.push(font.clone());
        self.fonts.len() as u32 - 1
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct Fills {
    #[serde(rename = "@count", default)]
    count: u32,
    #[serde(rename = "fill", default)]
    fills: Vec<Fill>
}

impl Fills {
    fn default() -> Fills {
        Fills {
            count: 0,
            fills: vec![],
        }
    }

    pub(crate) fn add_fill(&mut self, fill: &Fill) -> u32 {
        for i in 0..self.fills.len() {
            if self.fills[i] == *fill {
                return i as u32;
            }
        }
        self.count += 1;
        self.fills.push(fill.clone());
        self.fills.len() as u32 - 1
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct Borders {
    #[serde(rename = "@count", default)]
    count: u32,
    border: Vec<Border>,
}

impl Borders {
    fn default() -> Borders {
        Borders {
            count: 0,
            border: vec![],
        }
    }

    pub(crate) fn add_border(&mut self, border: &Border) -> u32 {
        for i in 0..self.border.len() {
            if self.border[i] == *border {
                return i as u32;
            }
        }
        self.count += 1;
        self.border.push(border.clone());
        self.border.len() as u32 - 1
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct CellStyleXfs {
    #[serde(rename = "@count", default)]
    count: u32,
    xf: Vec<Xf>
}

#[derive(Debug, Deserialize, Serialize)]
struct CellXfs {
    #[serde(rename = "@count", default)]
    count: u32,
    xf: Vec<Xf>
}

impl CellXfs {
    pub(crate) fn add_xf(&mut self, xf: &Xf) -> u32 {
        for i in 0..self.xf.len() {
            if self.xf[i] == *xf {
                return i as u32;
            }
        }
        self.count += 1;
        self.xf.push(xf.clone());
        self.xf.len() as u32 - 1
    }
}

#[derive(Debug, Deserialize, Serialize, PartialEq, Clone)]
struct Xf {
    #[serde(rename = "@numFmtId", default)]
    num_fmt_id: u32,
    #[serde(rename = "@fontId", default)]
    font_id: u32,
    #[serde(rename = "@fillId", default)]
    fill_id: u32,
    #[serde(rename = "@borderId", default)]
    border_id: u32,
    #[serde(rename = "@xfId", skip_serializing_if = "Option::is_none")]
    xf_id: Option<u32>,
    #[serde(rename = "@applyFont", skip_serializing_if = "Option::is_none")]
    apply_font: Option<u32>,
}

impl Xf {
    fn default() -> Xf {
        Xf {
            num_fmt_id: 0,
            font_id: 0,
            fill_id: 0,
            border_id: 0,
            xf_id: Some(0),
            apply_font: None,
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct CellStyles {
    #[serde(rename = "@count", default)]
    count: u32,
    #[serde(rename = "cellStyle", default)]
    cell_styles: Vec<CellStyle>
}

#[derive(Debug, Deserialize, Serialize)]
struct CellStyle {
    #[serde(rename = "@name")]
    name: String,
    #[serde(rename = "@xfId", default)]
    xf_id: u32,
    #[serde(rename = "@builtinId", default)]
    builtin_id: u32,
}

#[derive(Debug, Deserialize, Serialize)]
struct Dxfs {
    #[serde(rename = "@count", default)]
    count: u32,
}

#[derive(Debug, Deserialize, Serialize)]
struct TableStyles {
    #[serde(rename = "@count", default)]
    count: u32,
    #[serde(rename = "@defaultTableStyle")]
    default_table_style: String,
    #[serde(rename = "@defaultPivotStyle")]
    default_pivot_style: String,
}

impl StyleSheet {
    pub(crate) fn add_format(&mut self, format: &Format) -> u32 {
        // update font format
        let font_id = match &format.font {
            Some(font) => self.fonts.add_font(font),
            None => 0
        };
        // update border format
        let border_id = match &format.border {
            Some(border) => self.borders.add_border(border),
            None => 0
        };
        // update fill format
        let fill_id = match &format.fill {
            Some(fill) => self.fills.add_fill(fill),
            None => 0
        };
        // update cell xfs and return the xf index
        let mut xf = Xf::default();
        xf.font_id = font_id;
        xf.border_id = border_id;
        xf.fill_id = fill_id;
        self.cell_xfs.add_xf(&xf)
    }
}

trait Rearrange<E: Clone + Eq + Hash> {
    fn distinct(elements: &Vec<E>) -> (Vec<E>, HashMap<usize, usize>) {
        let mut distinct_elements = HashMap::new();
        for i in 0..elements.len() {
            let e = &elements[i];
            if !distinct_elements.contains_key(e) {
                distinct_elements.insert(e, Vec::new());
            }
            distinct_elements.get_mut(e).unwrap().push(i);
        }
        let mut index_map = HashMap::new();
        let distinct_elements: Vec<(&E, &Vec<usize>)> = distinct_elements
            .iter()
            .map(|(&e, ids)| (e, ids))
            .collect();
        for i in 0..distinct_elements.len() {
            let (e, ids) = distinct_elements[i];
            ids.iter().for_each(|&id| { index_map.insert(id, i); });
        }
        // let distinct_elements: Vec<E> = distinct_elements.iter().map(|&e| e.0.clone()).collect();
        let elements = distinct_elements.iter().map(|&e| e.0.clone()).collect::<Vec<E>>();
        (elements, index_map)
    }
}

impl XmlIo<StyleSheet> for StyleSheet {
    fn from_path<P: AsRef<Path>>(file_path: P) -> StyleSheet {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::StylesFile).unwrap();
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let style_sheet = de::from_str(&xml).unwrap();
        style_sheet
    }

    fn save<P: AsRef<Path>>(&mut self, file_path: P) {
        let xml = se::to_string_with_root("styleSheet", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::StylesFile).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}