//!
//! This module contains the [`Format`] struct, which used to edit the style of the [`Cell`],
//! also defines some methods for working with [`Format`].
//!
pub use align::{FormatAlignType, FormatAlign};
pub use border::{FormatBorder, FormatBorderElement, FormatBorderType};
pub use color::FormatColor;
pub use fill::FormatFill;
pub use font::FormatFont;
use crate::Cell;

mod align;
mod color;
mod fill;
mod font;
mod border;

///
/// [`Format`] struct, which used to edit the style of the [`Cell`].
/// # Fields
/// | field        | type        | meaning                                                      |
/// | ------------ | ----------- | ------------------------------------------------------------ |
/// | `font`     | [`FormatFont`] | The [`Cell`]'s font formats |
/// | `border`    | [`FormatBorder`] | The [`Cell`]'s border formats |
/// | `fill`      | [`FormatFill`] | The [`Cell`]'s fill(background) formats |
/// | `align`   | [`FormatAlign`] | The [`Cell`]'s align formats |
#[derive(Default, Clone, Debug, PartialEq)]
pub struct Format {
    pub font: FormatFont,
    pub border: FormatBorder,
    pub fill: FormatFill,
    pub align: FormatAlign,
}

impl Format {
    /// Checks whether the font is bold, learn more about it in [`FormatFont`].
    ///
    /// ## Returns
    ///
    /// Returns `true` if the font is bold, `false` otherwise.
    pub fn is_bold(&self) -> bool {
        self.font.bold
    }
    /// Checks whether the font is italic, learn more about it in [`FormatFont`].
    ///
    /// ## Returns
    ///
    /// Returns `true` if the font is italic, `false` otherwise.

    pub fn is_italic(&self) -> bool {
        self.font.italic
    }
    /// Checks whether the font is underline, learn more about it in [`FormatFont`].
    ///
    /// ## Returns
    ///
    /// Returns `true` if the font is underline, `false` otherwise.

    pub fn is_underline(&self) -> bool {
        self.font.underline
    }
    /// Retrieves the size of the font, learn more about it in [`FormatFont`].
    ///
    /// ## Returns
    ///
    /// Returns the pounds of the font size.
    pub fn get_size(&self) -> f64 {
        self.font.size
    }
    /// Retrieves the background(fill) of the cell.
    ///
    /// ## Returns
    ///
    /// Returns the [`FormatFill`] of the cell.
    pub fn get_background(&self) -> &FormatFill {
        &self.fill
    }

    /// Retrieves the borders of the cell, learn more about it in [`FormatBorder`].
    ///
    /// ## Returns
    ///
    /// Returns the [`FormatBorder`] of the cell.
    pub fn get_borders(&self) -> &FormatBorder {
        &self.border
    }
    /// Retrieves the color of the font, learn more about it in [`FormatFont`].
    ///
    /// ## Returns
    ///
    /// Returns the [`FormatColor`] of the font.
    pub fn get_color(&self) -> &FormatColor {
        &self.font.color
    }
    /// Retrieves the name of the font, learn more about it in [`FormatFont`].
    ///
    /// ## Returns
    ///
    /// Returns the name of the font.
    pub fn get_font(&self) -> &str {
        &self.font.name
    }

    /// Retrieves the color of the cell's background.
    ///
    /// ## Returns
    ///
    /// Returns the [`FormatColor`] of the cell's background.
    pub fn get_background_color(&self) -> &FormatColor {
        &self.fill.fg_color
    }

    /// Retrieves the left border of the cell, learn more about it in [`FormatBorder`].
    ///
    /// ## Returns
    ///
    /// Returns the [`FormatBorderElement`] of the cell's left border.
    pub fn get_border_left(&self) -> &FormatBorderElement {
        &self.border.left
    }
    /// Retrieves the right border of the cell, learn more about it in [`FormatBorder`].
    ///
    /// ## Returns
    ///
    /// Returns the [`FormatBorderElement`] of the cell's right border.
    pub fn get_border_right(&self) -> &FormatBorderElement {
        &self.border.right
    }
    /// Retrieves the top border of the cell, learn more about it in [`FormatBorder`].
    ///
    /// ## Returns
    ///
    /// Returns the [`FormatBorderElement`] of the cell's top border.
    pub fn get_border_top(&self) -> &FormatBorderElement {
        &self.border.top
    }
    /// Retrieves the bottom border of the cell, learn more about it in [`FormatBorder`].
    ///
    /// ## Returns
    ///
    /// Returns the [`FormatBorderElement`] of the cell's bottom border.
    pub fn get_border_bottom(&self) -> &FormatBorderElement {
        &self.border.bottom
    }
}

impl Format {
    /// Set the font to bold, learn more about it in [`FormatFont`].
    ///
    /// ## Returns
    ///
    /// Returns the modified format with font bold set to true.
    pub fn set_bold(mut self) -> Self {
        self.font.bold = true;
        self
    }
    /// Set the font to italic, learn more about it in [`FormatFont`].
    ///
    /// ## Returns
    ///
    /// Returns the modified format with font italic set to true.
    pub fn set_italic(mut self) -> Self {
        self.font.italic = true;
        self
    }
    /// Add an underline for the font, learn more about it in [`FormatFont`].
    ///
    /// ## Returns
    ///
    /// Returns the modified format with font underline set to true.
    pub fn set_underline(mut self) -> Self {
        self.font.underline = true;
        self
    }
    /// Set the size of the font, learn more about it in [`FormatFont`].
    ///
    /// ## Arguments
    ///
    /// | arg | type | meaning                                       |
    /// |----------|------|-----------------------------------------------|
    /// | size     | u8   | The new size of the font. |
    ///
    /// ## Returns
    ///
    /// Returns the modified Format with the size set to the provided value.
    pub fn set_size(mut self, size: u8) -> Self {
        self.font.size = size as f64;
        self
    }

    /// Set the size of the font, learn more about it in [`FormatFont`].
    ///
    /// ## Arguments
    ///
    /// | arg | type | meaning                                       |
    /// |----------|------|-----------------------------------------------|
    /// | size     | f64   | The new size of the font. |
    ///
    /// ## Returns
    ///
    /// Returns the modified Format with the size set to the provided value.
    pub fn set_size_f64(mut self, size: f64) -> Self {
        self.font.size = size;
        self
    }

    /// Set the color of the font, learn more about it in [`FormatFont`].
    ///
    /// ## Arguments
    ///
    /// | arg | type | meaning                                       |
    /// |----------|------|-----------------------------------------------|
    /// | format_color     | [`FormatColor`]   | The new color of the font. |
    ///
    /// ## Returns
    ///
    /// Returns the modified Format with the color set to the provided value.
    pub fn set_color(mut self, format_color: FormatColor) -> Self {
        self.font.color = format_color;
        self
    }

    /// Set the font name, learn more about it in [`FormatFont`].
    ///
    /// ## Arguments
    ///
    /// | arg | type | meaning                                       |
    /// |----------|------|-----------------------------------------------|
    /// | size     | &str   | The new font name. |
    ///
    /// ## Returns
    ///
    /// Returns the modified Format with the name set to the provided value.
    pub fn set_font(mut self, font_name: &str) -> Self {
        self.font.name = font_name.to_string();
        self
    }

    /// Set all borders of the format, learn more about it in [`FormatBorder`].
    ///
    /// ## Arguments
    ///
    /// | arg            | type                  | meaning                                       |
    /// |---------------------|-----------------------|-----------------------------------------------|
    /// | format_border_type | [`FormatBorderType`]      | The type of border to apply to borders of the format. |
    ///
    /// ## Returns
    ///
    /// Returns the modified Format with borders of the format set to the provided type.
    pub fn set_border(mut self, format_border_type: FormatBorderType) -> Self {
        let mut format_border = FormatBorderElement::default();
        format_border.border_type = format_border_type;
        self.border.left = format_border.clone();
        self.border.right = format_border.clone();
        self.border.top = format_border.clone();
        self.border.bottom = format_border.clone();
        self.border.diagonal = format_border;
        self
    }

    /// Set the left border of the format, learn more about it in [`FormatBorder`].
    ///
    /// ## Arguments
    ///
    /// | arg            | type                  | meaning                                       |
    /// |---------------------|-----------------------|-----------------------------------------------|
    /// | format_border_type | [`FormatBorderType`]      | The type of border to apply to the left side of the format. |
    ///
    /// ## Returns
    ///
    /// Returns the modified Format with the left border of the format set to the provided type.
    pub fn set_border_left(mut self, format_border_type: FormatBorderType) -> Self {
        let mut format_border = FormatBorderElement::default();
        format_border.border_type = format_border_type;
        self.border.left = format_border;
        self
    }

    /// Set the right border of the format, learn more about it in [`FormatBorder`].
    ///
    /// ## Arguments
    ///
    /// | arg            | type                  | meaning                                       |
    /// |---------------------|-----------------------|-----------------------------------------------|
    /// | format_border_type | [`FormatBorderType`]      | The type of border to apply to the right side of the format. |
    ///
    /// ## Returns
    ///
    /// Returns the modified Format with the right border of the format set to the provided type.
    pub fn set_border_right(mut self, format_border_type: FormatBorderType) -> Self {
        let mut format_border = FormatBorderElement::default();
        format_border.border_type = format_border_type;
        self.border.right = format_border;
        self
    }

    /// Set the top border of the format, learn more about it in [`FormatBorder`].
    ///
    /// ## Arguments
    ///
    /// | arg            | type                  | meaning                                       |
    /// |---------------------|-----------------------|-----------------------------------------------|
    /// | format_border_type | [`FormatBorderType`]      | The type of border to apply to the top side of the format. |
    ///
    /// ## Returns
    ///
    /// Returns the modified Format with the top border of the format set to the provided type.
    pub fn set_border_top(mut self, format_border_type: FormatBorderType) -> Self {
        let mut format_border = FormatBorderElement::default();
        format_border.border_type = format_border_type;
        self.border.top = format_border;
        self
    }
    /// Set the bottom border of the format, learn more about it in [`FormatBorder`].
    ///
    /// ## Arguments
    ///
    /// | arg            | type                  | meaning                                       |
    /// |---------------------|-----------------------|-----------------------------------------------|
    /// | format_border_type | [`FormatBorderType`]      | The type of border to apply to the bottom side of the format. |
    ///
    /// ## Returns
    ///
    /// Returns the modified Format with the bottom border of the format set to the provided type.
    pub fn set_border_bottom(mut self, format_border_type: FormatBorderType) -> Self {
        let mut format_border = FormatBorderElement::default();
        format_border.border_type = format_border_type;
        self.border.bottom = format_border;
        self
    }
    
    /// Set the background color of the format.
    ///
    /// ## Arguments
    ///
    /// | arg       | type         | meaning                                       |
    /// |----------------|--------------|-----------------------------------------------|
    /// | format_color   | [`FormatColor`]  | The color to set as the background of the format. |
    ///
    /// ## Returns
    ///
    /// Returns the modified Format with the background color of the format set to the provided color.
    pub fn set_background_color(mut self, format_color: FormatColor) -> Self {
        self.fill.pattern_type = "solid".to_string();
        self.fill.fg_color = format_color;
        self
    }

    /// Set the alignment of the format, learn more about it in [`FormatAlign`].
    ///
    /// ## Arguments
    ///
    /// | arg             | type                  | meaning                                       |
    /// |----------------------|-----------------------|-----------------------------------------------|
    /// | format_align_type    | [`FormatAlignType`]       | The type of alignment to apply to the format. |
    ///
    /// ## Returns
    ///
    /// Returns the modified Format with the alignment of the format set to the provided type.
    pub fn set_align(mut self, format_align_type: FormatAlignType) -> Self {
        match format_align_type {
            FormatAlignType::Left | FormatAlignType::Center | FormatAlignType::Right =>
                self.align.horizontal = Some(format_align_type),
            FormatAlignType::Top | FormatAlignType::VerticalCenter | FormatAlignType::Bottom =>
                self.align.vertical = Some(format_align_type),
        }
        self
    }

    /// Set the indent of the text alignment, learn more about it in [`FormatAlign`].
    ///
    /// ## Arguments
    ///
    /// | arg | type | meaning                                       |
    /// |----------|------|-----------------------------------------------|
    /// | indent   | u8   | The new reading order of the text alignment. |
    ///
    /// ## Returns
    ///
    /// Returns the modified Format with the reading order of the text alignment set to the provided value.
    pub fn set_reading_order(mut self, reading_order: u8) -> Self {
        self.align.reading_order = reading_order;
        self
    }

    /// Set the indent of the text alignment, learn more about it in [`FormatAlign`].
    ///
    /// ## Arguments
    ///
    /// | arg | type | meaning                                       |
    /// |----------|------|-----------------------------------------------|
    /// | indent   | u8   | The new indent of the text alignment. |
    ///
    /// ## Returns
    ///
    /// Returns the modified Format with the indent of the text alignment set to the provided value.
    pub fn set_indent(mut self, indent: u8) -> Self {
        self.align.indent = indent;
        self
    }
}