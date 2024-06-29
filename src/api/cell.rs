use crate::api::cell::formula::Formula;
use crate::api::cell::rich_text::RichText;
use crate::api::cell::values::{CellDisplay, CellType, CellValue};
use crate::{Format, FormatColor, FormatFont};

pub mod formula;
pub mod location;
pub mod values;
pub mod rich_text;

#[derive(Clone, Debug, Default)]
pub struct Cell<T: CellDisplay + CellValue> {
    pub text: Option<T>,
    pub rich_text: Option<RichText>,
    pub format: Option<Format>,
    pub hyperlink: Option<String>,
    pub(crate) formula: Option<Formula>,
    pub(crate) cell_type: Option<CellType>,
    pub(crate) style: Option<u32>,
}

// use ansi_term::{ANSIString, Style, Colour};
// impl<T: CellDisplay + CellValue> Cell<T> {
//
//     pub fn ansi_strings(&self) -> Vec::<ANSIString> {
//         fn to_ansi_style(format_font: &FormatFont) -> Style {
//             let mut style = Style::new();
//             style.is_bold = format_font.bold;
//             style.is_italic = format_font.italic;
//             style.is_underline = format_font.underline;
//             match format_font.color {
//                 FormatColor::RGB(r, g, b) => {
//                     style.foreground = Some(Colour::RGB(r, g, b));
//                 }
//                 _ => {}
//             }
//             style
//         }
//
//         let default_style = if let Some(format) = &self.format {
//             let mut default_style = to_ansi_style(&format.font);
//             match format.fill.fg_color {
//                 FormatColor::RGB(r, g, b) => {
//                     default_style.background = Some(Colour::RGB(r, g, b));
//                 }
//                 _ => {}
//             }
//             default_style
//         } else {
//             Style::new()
//         };
//         let mut ansi_string_vec = Vec::new();
//         let text = if let Some(text) = &self.text {
//             text.to_display()
//         } else { "".to_string() };
//         if let Some(rich_text) = &self.rich_text {
//             rich_text.words.iter().for_each(|w| {
//                 let mut style = default_style.clone();
//                 if let Some(format_font) = &w.font {
//                     style = to_ansi_style(format_font);
//                 }
//                 ansi_string_vec.push(
//                     style.paint(w.text.clone())
//                 );
//             });
//         } else {
//             ansi_string_vec.push(
//                 default_style.paint(text)
//             );
//         };
//         ansi_string_vec
//     }
// }
