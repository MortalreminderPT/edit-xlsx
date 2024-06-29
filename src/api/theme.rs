use crate::{Cell, FormatColor, FormatFont};

pub struct Theme {
    colors_rgb: Vec<FormatColor>,
}

impl Theme {
    pub(crate) fn new(colors_rgb: Vec<FormatColor>) -> Theme {
        Theme {
            colors_rgb
        }
    }

    pub(crate) fn theme_to_rgb(&self, color_theme: u8, tint: f64) -> &FormatColor {
        self.colors_rgb.get(color_theme as usize).unwrap()
    }

    pub(crate) fn index_to_rgb(&self, index: u8) -> FormatColor {
        let (r, g, b) = match index {
            0 => (0, 0, 0),          // Black
            1 => (255, 255, 255),    // White
            2 => (255, 0, 0),        // Red
            3 => (0, 255, 0),        // Bright Green
            4 => (0, 0, 255),        // Blue
            5 => (255, 255, 0),      // Yellow
            6 => (255, 0, 255),      // Pink
            7 => (0, 255, 255),      // Turquoise
            8 => (0, 0, 0),          // Black
            9 => (255, 255, 255),    // White
            10 => (255, 0, 0),       // Red
            11 => (0, 255, 0),       // Bright Green
            12 => (0, 0, 255),       // Blue
            13 => (255, 255, 0),     // Yellow
            14 => (255, 0, 255),     // Pink
            15 => (0, 255, 255),     // Turquoise
            16 => (128, 0, 0),       // Dark Red
            17 => (0, 128, 0),       // Green
            18 => (0, 0, 128),       // Dark Blue
            19 => (128, 128, 0),     // Dark Yellow
            20 => (128, 0, 128),     // Violet
            21 => (0, 128, 128),     // Teal
            22 => (192, 192, 192),   // Gray-25%
            23 => (128, 128, 128),   // Gray-50%
            24 => (153, 153, 255),   // Periwinkle
            25 => (153, 51, 102),    // Plum
            26 => (255, 255, 204),   // Ivory
            27 => (204, 255, 255),   // Light Turquoise
            28 => (102, 0, 102),     // Dark Purple
            29 => (255, 128, 128),   // Coral
            30 => (0, 102, 204),     // Ocean Blue
            31 => (204, 204, 255),   // Ice Blue
            32 => (0, 0, 128),       // Dark Blue
            33 => (255, 0, 255),     // Pink
            34 => (255, 255, 0),     // Yellow
            35 => (0, 255, 255),     // Turquoise
            36 => (128, 0, 128),     // Violet
            37 => (128, 0, 0),       // Dark Red
            38 => (0, 128, 128),     // Teal
            39 => (0, 0, 255),       // Blue
            40 => (0, 204, 255),     // Sky Blue
            41 => (204, 255, 255),   // Light Turquoise
            42 => (204, 255, 204),   // Light Green
            43 => (255, 255, 153),   // Light Yellow
            44 => (153, 204, 255),   // Pale Blue
            45 => (255, 153, 204),   // Rose
            46 => (204, 153, 255),   // Lavender
            47 => (255, 204, 153),   // Tan
            48 => (51, 102, 255),    // Light Blue
            49 => (51, 204, 204),    // Aqua
            50 => (153, 204, 0),     // Lime
            51 => (255, 204, 0),     // Gold
            52 => (255, 153, 0),     // Light Orange
            53 => (255, 102, 0),     // Orange
            54 => (102, 102, 153),   // Blue-Gray
            55 => (150, 150, 150),   // Gray-Gray40%
            56 => (0, 51, 102),      // Dark Teal
            57 => (51, 153, 102),    // Sea Green
            58 => (0, 51, 0),        // Dark Green
            59 => (51, 51, 0),       // Olive Green
            60 => (153, 51, 0),      // Brown
            61 => (153, 51, 102),    // Plum
            62 => (51, 51, 153),     // Indigo
            63 => (51, 51, 51),      // Gray-80%
            _ => (0, 0, 0),          // Default to Black for unspecified indices
        };
        FormatColor::RGB(r, g, b)
    }
}

// use ansi_term::{ANSIString, Style, Colour};
// use crate::api::cell::values::{CellDisplay, CellValue};

// impl Theme {
//     pub fn ansi_strings<T: CellDisplay + CellValue>(&self, cell: &Cell<T>) -> Vec::<ANSIString> {
//         fn to_ansi_style(theme: &Theme, format_font: &FormatFont) -> Style {
//             let mut style = Style::new();
//             style.is_bold = format_font.bold;
//             style.is_italic = format_font.italic;
//             style.is_underline = format_font.underline;
//             let mut color = format_font.color;
//             if let FormatColor::Theme(id, tint) = format_font.color {
//                 color = *theme.theme_to_rgb(id, tint);
//             }
//             style.foreground = color.to_ansi_term_colour();
//             style
//         }
//
//         let default_style = if let Some(format) = &cell.format {
//             let mut default_style = to_ansi_style(&self, &format.font);
//             let mut color = format.fill.fg_color;
//             if let FormatColor::Theme(id, tint) = format.fill.fg_color {
//                 color = *self.theme_to_rgb(id, tint);
//             }
//             default_style.background = color.to_ansi_term_colour();
//             default_style
//         } else {
//             Style::new()
//         };
//         let mut ansi_string_vec = Vec::new();
//         let text = if let Some(text) = &cell.text {
//             text.to_display()
//         } else { "".to_string() };
//         if let Some(rich_text) = &cell.rich_text {
//             rich_text.words.iter().for_each(|w| {
//                 let mut style = default_style.clone();
//                 if let Some(format_font) = &w.font {
//                     style = to_ansi_style(&self, &format_font);
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

