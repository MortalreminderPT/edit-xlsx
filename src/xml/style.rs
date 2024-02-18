pub(crate) mod font;
pub(crate) mod border;
pub(crate) mod fill;
pub(crate) mod alignment;
pub(crate) mod xf;
pub(crate) mod color;

use std::collections::HashMap;
use std::hash::Hash;
use std::io;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::api::format::Format;
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::common::{ExtLst, FromFormat, XmlnsAttrs};
use crate::xml::manage::Io;
use crate::xml::style::alignment::Alignment;
use crate::xml::style::border::{Border, Borders};
use crate::xml::style::fill::{Fill, Fills};
use crate::xml::style::font::{Font, Fonts};
use crate::xml::style::xf::Xf;

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
struct CellStyleXfs {
    #[serde(rename = "@count", default)]
    count: u32,
    xf: Vec<Xf>
}

impl Default for CellStyleXfs {
    fn default() -> Self {
        CellStyleXfs {
            count: 1,
            xf: vec![Xf::default()],
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct CellXfs {
    #[serde(rename = "@count", default)]
    count: u32,
    xf: Vec<Xf>
}

impl Default for CellXfs {
    fn default() -> Self {
        CellXfs {
            count: 1,
            xf: vec![Xf::default()],
        }
    }
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

#[derive(Debug, Deserialize, Serialize)]
struct CellStyles {
    #[serde(rename = "@count", default)]
    count: u32,
    #[serde(rename = "cellStyle", default)]
    cell_styles: Vec<CellStyle>
}

impl Default for CellStyles {
    fn default() -> Self {
        CellStyles {
            count: 1,
            cell_styles: vec![Default::default()],
        }
    }
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

impl Default for CellStyle {
    fn default() -> Self {
        CellStyle {
            name: "Normal".to_string(),
            xf_id: 0,
            builtin_id: 0,
        }
    }
}

#[derive(Debug, Deserialize, Serialize, Default)]
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

impl Default for TableStyles {
    fn default() -> Self {
        TableStyles {
            count: 1,
            default_table_style: "TableStyleMedium2".to_string(),
            default_pivot_style: "PivotStyleLight16".to_string(),
        }
    }
}

impl Default for StyleSheet {
    fn default() -> Self {
        StyleSheet {
            xmlns_attrs: XmlnsAttrs::stylesheet_default(),
            fonts: Default::default(),
            fills: Default::default(),
            borders: Default::default(),
            cell_style_xfs: CellStyleXfs::default(),
            cell_xfs: Default::default(),
            cell_styles: Default::default(),
            dxfs: Dxfs::default(),
            table_styles: Default::default(),
            ext_lst: None,
        }
    }
}

impl StyleSheet {
    pub(crate) fn add_format(&mut self, format: &Format) -> u32 {
        let font = Font::from_format(&format.font);
        let font_id = self.fonts.add_font(&font);
        let border = Border::from_format(&format.border);
        let border_id = self.borders.add_border(&border);
        let fill = Fill::from_format(&format.fill);
        let fill_id = self.fills.add_fill(&fill);
        let mut xf = Xf::default();
        let align = Alignment::from_format(&format.align);
        xf.alignment = Some(align);
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
            let (_, ids) = distinct_elements[i];
            ids.iter().for_each(|&id| { index_map.insert(id, i); });
        }
        // let distinct_elements: Vec<E> = distinct_elements.iter().map(|&e| e.0.clone()).collect();
        let elements = distinct_elements.iter().map(|&e| e.0.clone()).collect::<Vec<E>>();
        (elements, index_map)
    }
}

impl Io<StyleSheet> for StyleSheet {
    fn from_path<P: AsRef<Path>>(file_path: P) -> io::Result<StyleSheet> {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::StylesFile)?;
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let style_sheet = de::from_str(&xml).unwrap();
        Ok(style_sheet)
    }

    fn save<P: AsRef<Path>>(&mut self, file_path: P) {
        let xml = se::to_string_with_root("styleSheet", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::StylesFile).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}