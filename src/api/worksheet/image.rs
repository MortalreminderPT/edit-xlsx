use std::path::Path;
use crate::api::cell::location::LocationRange;
use crate::api::worksheet::WorkSheet;
use crate::xml::drawings::Drawings;
use crate::xml::relationships::Relationships;

pub(crate) trait _Image {
    fn add_background<P: AsRef<Path>>(&mut self, filename: &P) -> u32;
    fn add_drawing<L: LocationRange,P: AsRef<Path>>(&mut self, loc: L, filename: &P) -> u32;
}

impl _Image for WorkSheet {
    fn add_background<P: AsRef<Path>>(&mut self, image_path: &P) -> u32 {
        self.content_types.borrow_mut().add_png();
        let image_id = self.medias.borrow_mut().add_media(image_path);
        self.worksheet_rel.add_image(image_id)
    }

    fn add_drawing<L: LocationRange, P: AsRef<Path>>(&mut self, loc: L, image_path: &P) -> u32 {
        // get drawings file
        let drawings = self.drawings.get_or_insert(Drawings::default());
        let drawings_rel = &mut self.drawings_rel.get_or_insert(Relationships::default());
        self.content_types.borrow_mut().add_png();
        self.content_types.borrow_mut().add_drawing(self.id);
        let image_id = self.medias.borrow_mut().add_media(image_path);
        let image_r_id = drawings_rel.add_image(image_id);
        drawings.add_drawing(loc, image_r_id);
        let r_id = self.worksheet_rel.add_drawings(self.id);
        r_id
    }
}