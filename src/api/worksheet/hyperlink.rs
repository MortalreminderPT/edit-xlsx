use crate::api::worksheet::WorkSheet;

pub(crate) trait _Hyperlink {
    fn add_hyperlink(&mut self, hyperlink: &str) -> u32;
}

impl _Hyperlink for WorkSheet {
    fn add_hyperlink(&mut self, hyperlink: &str) -> u32 {
        self.worksheet_rel.add_hyperlink(hyperlink)
    }
}