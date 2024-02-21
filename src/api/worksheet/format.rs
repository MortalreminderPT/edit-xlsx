use crate::api::worksheet::WorkSheet;
use crate::Format;

pub(crate) trait _Format {
    fn add_format(&mut self, format: &Format) -> u32;
}

impl _Format for WorkSheet {
    fn add_format(&mut self, format: &Format) -> u32 {
        self.style_sheet.borrow_mut().add_format(format)
    }
}