pub(crate) mod font;
pub(crate) mod border;

use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::Format;
use crate::xml::common::{ExtLst, XmlnsAttrs};
use crate::xml::manage::XmlIo;
use crate::xml::style::border::Border;
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

#[derive(Debug, Deserialize, Serialize)]
struct Fills {
    #[serde(rename = "@count", default)]
    count: u32,
    #[serde(rename = "fill", default)]
    fills: Vec<Fill>
}

#[derive(Debug, Deserialize, Serialize)]
struct Fill {
    #[serde(rename = "patternFill")]
    pattern_fill: PatternFill
}

#[derive(Debug, Deserialize, Serialize)]
struct PatternFill {
    #[serde(rename = "@patternType")]
    pattern_type: String,
}

#[derive(Debug, Deserialize, Serialize)]
struct Borders {
    #[serde(rename = "@count", default)]
    count: u32,
    border: Vec<Border>
}

impl Borders {
    fn default() -> Borders {
        Borders {
            count: 0,
            border: vec![],
        }
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

#[derive(Debug, Deserialize, Serialize)]
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
    pub(crate) fn add_format(&mut self, format: Format) -> u32 {
        // update font format
        let font_id = if let Some(font) = format.font {
            self.fonts.fonts.push(font);
            self.fonts.count += 1;
            self.fonts.count - 1
        } else { 0 };
        // update border format
        let border_id = if let Some(border) = format.border {
            self.borders.border.push(border);
            self.borders.count += 1;
            self.borders.count - 1
        } else { 0 };
        // update cell xfs and return the xf index
        let mut xf = Xf::default();
        xf.font_id = font_id;
        xf.border_id = border_id;
        self.cell_xfs.xf.push(xf);
        self.cell_xfs.count += 1;
        self.cell_xfs.count - 1
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