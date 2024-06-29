use crate::api::worksheet::WorkSheet;

#[cfg(feature = "ansi_term_support")]
use ansi_term::{ANSIString, Style, Colour};
use crate::api::cell::location::{Location, LocationRange};
use crate::api::cell::values::CellDisplay;
use crate::api::theme::Theme;
use crate::{FormatColor, FormatFill, FormatFont, Read, WorkSheetCol, WorkSheetResult, WorkSheetRow};

#[cfg(feature = "ansi_term_support")]
impl WorkSheet {
    pub fn ansi_strings<L: Location>(&self, loc: L) -> WorkSheetResult<Vec::<ANSIString>> {
        fn add_ansi_font_style(theme: &Theme, format_font: &FormatFont, style: &mut Style) {
            style.is_bold |= format_font.bold;
            style.is_italic |= format_font.italic;
            style.is_underline |= format_font.underline;
            let mut color = format_font.color;
            if let FormatColor::Theme(id, tint) = format_font.color {
                color = *theme.theme_to_rgb(id, tint);
            }
            if let FormatColor::Index(id) = format_font.color {
                color = theme.index_to_rgb(id);
            }
            if let Some(color) = color.to_ansi_term_colour() {
                style.foreground = Some(color)
            }
        }
        fn add_ansi_color_style(theme: &Theme, format_fill: &FormatFill, style: &mut Style) {
            let mut color = format_fill.fg_color;
            if let FormatColor::Theme(id, tint) = format_fill.fg_color {
                color = *theme.theme_to_rgb(id, tint);
            }
            if let FormatColor::Index(id) = format_fill.fg_color {
                color = theme.index_to_rgb(id);
            }
            if let Some(color) = color.to_ansi_term_colour() {
                style.background = Some(color)
            }
        }

        let (row, col) = loc.to_location();
        let theme = self.get_theme(0);
        let mut style = Style::new();

        let (_, format) = self.get_row_with_format(row)?;
        if let Some(format) = format {
            add_ansi_font_style(&theme, &format.font, &mut style);
            add_ansi_color_style(&theme, &format.fill, &mut style);
        }

        let col_range = (row, col, row, col).to_col_range_ref();
        let col_map =
            self.get_columns_with_format(&col_range)?;
        if let Some((_, Some(format))) = col_map.get(&col_range) {
            add_ansi_font_style(&theme, &format.font, &mut style);
            add_ansi_color_style(&theme, &format.fill, &mut style);
        }

        let mut ansi_string_vec = Vec::new();
        if let Ok(cell) = self.read_cell((row, col)) {
            if let Some(format) = cell.format {
                add_ansi_font_style(&theme, &format.font, &mut style);
                add_ansi_color_style(&theme, &format.fill, &mut style);
            }
            let text = if let Some(text) = &cell.text {
                text.to_display()
            } else { "".to_string() };
            if let Some(rich_text) = &cell.rich_text {
                rich_text.words.iter().for_each(|w| {
                    let mut style = style;
                    if let Some(format_font) = &w.font {
                        add_ansi_font_style(&theme, &format_font, &mut style);
                    }
                    ansi_string_vec.push(
                        style.paint(w.text.clone())
                    );
                });
            } else {
                ansi_string_vec.push(
                    style.paint(text)
                );
            }
        }
        Ok(ansi_string_vec)
    }
}

#[cfg(feature = "ansi_term_support")]
impl FormatColor {
    pub(crate) fn to_ansi_term_colour(&self) -> Option<Colour> {
        match self {
            FormatColor::RGB(r, g, b) => { Some(Colour::RGB(*r, *g, *b)) }
            _ => { None }
        }
    }
}