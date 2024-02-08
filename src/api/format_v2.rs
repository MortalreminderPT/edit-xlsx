use crate::{FormatAlign, FormatColor};
use crate::api::fill::FormatFill;
use crate::api::font::FormatFont;
use crate::FormatBorder as FormatBorderType;

struct Format<'a> {
    font: FormatFont<'a>,
    border: FormatBorder,
    fill: FormatFill<'a>,
    align: FormatAlign,
}

struct FormatBorder {
    left: FormatBorderType,
    right: FormatBorderType,
    top: FormatBorderType,
    bottom: FormatBorderType,
    diagonal: FormatBorderType,
}
impl Default for Format {
    fn default() -> Self {
        Format {
            font: Default::default(),
            border: FormatBorder::None,
            fill: Default::default(),
            align: FormatAlign::Top,
        }
    }
}