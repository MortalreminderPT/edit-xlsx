use std::path::Path;
use crate::api::worksheet::Sheet;
use crate::xml::worksheet_rel::Relationships;

pub(crate) trait _Image {
    fn add_image<P: AsRef<Path>>(&self, filename: &P) -> u32;
}

impl _Image for Sheet {
    fn add_image<P: AsRef<Path>>(&self, filename: &P) -> u32 {
        let mut worksheets_rel = self.worksheets_rel.borrow_mut();
        if let None = worksheets_rel.get_mut(&self.id) {
            worksheets_rel.insert(self.id, Relationships::default());
        }
        let worksheet_rel = worksheets_rel.get_mut(&self.id).unwrap();
        self.content_types.borrow_mut().add_png();
        let image_id = self.medias.borrow_mut().add_media(filename);
        let r_id = worksheet_rel.next_id();
        worksheet_rel.add_image(r_id, image_id);
        r_id
    }
}