use crate::api::worksheet::WorkSheet;
use crate::Format;

pub(crate) trait _Format {
    fn add_format(&mut self, format: &Format) -> u32;
    fn get_format(&self, style_id: u32) -> Format;
}

impl _Format for WorkSheet {
    fn add_format(&mut self, format: &Format) -> u32 {
        self.style_sheet.borrow_mut().add_format(format)
    }

    fn get_format(&self, style_id: u32) -> Format {
        let mut format = Format::default();
        self.style_sheet.borrow().update_format(&mut format, style_id);
        format
    }
}