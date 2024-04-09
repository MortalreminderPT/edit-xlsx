pub(crate) mod font;
pub(crate) mod border;
pub(crate) mod fill;
pub(crate) mod alignment;
pub(crate) mod xf;
pub(crate) mod color;
mod num_fmt;

use std::hash::Hash;
use std::io;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::api::format::{Format, FormatFont};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::xml::common::{FromFormat, XmlnsAttrs};
use crate::xml::extension::ExtensionList;
use crate::xml::io::Io;
use crate::xml::style::alignment::Alignment;
use crate::xml::style::border::{Border, Borders};
use crate::xml::style::fill::{Fill, Fills};
use crate::xml::style::font::{Font, Fonts};
use crate::xml::style::num_fmt::NumFmts;
use crate::xml::style::xf::Xf;

#[derive(Debug, Clone, Deserialize, Serialize)]
pub(crate) struct StyleSheet {
    #[serde(flatten)]
    xmlns_attrs: XmlnsAttrs,
    #[serde(rename = "numFmts", skip_serializing_if = "Option::is_none")]
    num_fmts: Option<NumFmts>,
    #[serde(rename = "fonts", skip_serializing_if = "Option::is_none")]
    pub(crate) fonts: Option<Fonts>,
    #[serde(rename = "fills", skip_serializing_if = "Option::is_none")]
    pub(crate) fills: Option<Fills>,
    #[serde(rename = "borders", skip_serializing_if = "Option::is_none")]
    pub(crate) borders: Option<Borders>,
    #[serde(rename = "cellStyleXfs", skip_serializing_if = "Option::is_none")]
    cell_style_xfs: Option<CellStyleXfs>,
    #[serde(rename = "cellXfs", skip_serializing_if = "Option::is_none")]
    pub(crate) cell_xfs: Option<CellXfs>,
    #[serde(rename = "cellStyles", skip_serializing_if = "Option::is_none")]
    cell_styles: Option<CellStyles>,
    #[serde(rename = "dxfs", skip_serializing_if = "Option::is_none")]
    dxfs: Option<Dxfs>,
    #[serde(rename = "tableStyles", skip_serializing_if = "Option::is_none")]
    table_styles: Option<TableStyles>,
    #[serde(rename = "extLst", skip_serializing_if = "Option::is_none")]
    ext_lst: Option<ExtensionList>,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
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

#[derive(Debug, Clone, Deserialize, Serialize)]
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

    pub(crate) fn get_xf(&self, id: u32) -> Option<&Xf> {
        self.xf.get(id as usize)
    }
}

#[derive(Debug, Clone, Deserialize, Serialize)]
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

#[derive(Debug, Clone, Deserialize, Serialize)]
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

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
struct Dxfs {
    #[serde(rename = "@count", default)]
    count: u32,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
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
            num_fmts: None,
            fonts: Default::default(),
            fills: Default::default(),
            borders: Default::default(),
            cell_style_xfs: None, //CellStyleXfs::default(),
            cell_xfs: Default::default(),
            cell_styles: Default::default(),
            dxfs: None,//Dxfs::default(),
            table_styles: Default::default(),
            ext_lst: None,
        }
    }
}

impl StyleSheet {
    pub(crate) fn add_format(&mut self, format: &Format) -> u32 {
        let fonts = self.fonts.get_or_insert(Default::default());
        let font = Font::from_format(&format.font);
        let font_id = fonts.add_font(&font);
        let borders = self.borders.get_or_insert(Default::default());
        let border = Border::from_format(&format.border);
        let border_id = borders.add_border(&border);
        let fills = self.fills.get_or_insert(Default::default());
        let fill = Fill::from_format(&format.fill);
        let fill_id = fills.add_fill(&fill);
        let mut xf = Xf::default();
        let align = Alignment::from_format(&format.align);
        xf.alignment = Some(align);
        xf.font_id = font_id;
        xf.border_id = border_id;
        xf.fill_id = fill_id;
        let cell_xfs = self.cell_xfs.get_or_insert(Default::default());
        cell_xfs.add_xf(&xf)
    }

    pub(crate) fn update_format(&self, format: &mut Format, style_id: u32) {
        if let Some(cell_xfs) = &self.cell_xfs {
            if let Some(xf) = cell_xfs.get_xf(style_id) {
                let font = &self.fonts.as_ref().unwrap().get_font(xf.font_id);
                format.font = font.get_format();
                let border = &self.borders.as_ref().unwrap().get_border(xf.border_id);
                format.border = border.get_format();
                let fill = &self.fills.as_ref().unwrap().get_fill(xf.fill_id);
                format.fill = fill.get_format();
                // let fill = &self.fills.as_ref().unwrap().get_fill(xf.fill_id);
            }
        }
    }
}

// trait Rearrange<E: Clone + Eq + Hash> {
//     fn distinct(elements: &Vec<E>) -> (Vec<E>, HashMap<usize, usize>) {
//         let mut distinct_elements = HashMap::new();
//         for i in 0..elements.len() {
//             let e = &elements[i];
//             if !distinct_elements.contains_key(e) {
//                 distinct_elements.insert(e, Vec::new());
//             }
//             distinct_elements.get_mut(e).unwrap().push(i);
//         }
//         let mut index_map = HashMap::new();
//         let distinct_elements: Vec<(&E, &Vec<usize>)> = distinct_elements
//             .iter()
//             .map(|(&e, ids)| (e, ids))
//             .collect();
//         for i in 0..distinct_elements.len() {
//             let (_, ids) = distinct_elements[i];
//             ids.iter().for_each(|&id| { index_map.insert(id, i); });
//         }
//         let elements = distinct_elements.iter().map(|&e| e.0.clone()).collect::<Vec<E>>();
//         (elements, index_map)
//     }
// }

impl Io<StyleSheet> for StyleSheet {
    fn from_path<P: AsRef<Path>>(file_path: P) -> io::Result<StyleSheet> {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::StylesFile)?;
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let style_sheet = de::from_str(&xml).unwrap();
        Ok(style_sheet)
    }

    fn save<P: AsRef<Path>>(&self, file_path: P) {
        let xml = se::to_string_with_root("styleSheet", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::StylesFile).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}