use crate::api::worksheet::Sheet;

pub(crate) trait _Hyperlink {
    fn add_hyperlink(&mut self, hyperlink: &str) -> u32;
}

impl _Hyperlink for Sheet {
    fn add_hyperlink(&mut self, hyperlink: &str) -> u32 {
        let r_id = self.worksheet_rel.next_id();
        self.worksheet_rel.add_hyperlink(r_id, hyperlink);
        r_id
    }
}