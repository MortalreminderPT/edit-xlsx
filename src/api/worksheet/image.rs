use std::path::Path;
use futures::io;
use crate::api::cell::location::LocationRange;
use crate::api::worksheet::WorkSheet;
use crate::result::WorkSheetError;
use crate::WorkSheetResult;
use crate::xml::drawings::Drawings;
use crate::xml::relationships::Relationships;

pub(crate) trait Image {
    fn add_background<P: AsRef<Path>>(&mut self, filename: &P) -> WorkSheetResult<u32> ;
    fn add_drawing<L: LocationRange,P: AsRef<Path>>(&mut self, loc: L, filename: &P) -> WorkSheetResult<u32>;
}

impl Image for WorkSheet {
    fn add_background<P: AsRef<Path>>(&mut self, image_path: &P) -> WorkSheetResult<u32> {
        let extension = get_extension(image_path)?;
        if extension != "png" {
            return Err(WorkSheetError::FormatError);
        }
        self.content_types.borrow_mut().add_png();
        let image_id = self.medias.borrow_mut().add_media(image_path);
        Ok(self.worksheet_rel.add_image(image_id, extension))
    }

    fn add_drawing<L: LocationRange, P: AsRef<Path>>(&mut self, loc: L, image_path: &P) -> WorkSheetResult<u32> {
        // get extension
        let extension = get_extension(image_path)?;
        // get drawings file
        let drawings = self.drawings.get_or_insert(Drawings::default());
        let drawings_rel = &mut self.drawings_rel.get_or_insert(Relationships::default());
        self.content_types.borrow_mut().add_bin(extension);
        self.content_types.borrow_mut().add_drawing(self.id);
        let image_id = self.medias.borrow_mut().add_media(image_path);
        let image_r_id = drawings_rel.add_image(image_id, extension);
        drawings.add_drawing(loc, image_r_id);
        let r_id = self.worksheet_rel.add_drawings(self.id);
        Ok(r_id)
    }
}

fn get_extension<P: AsRef<Path>>(image_path: &P) -> WorkSheetResult<&str> {
    image_path.as_ref().extension().unwrap().to_str()
        .ok_or(WorkSheetError::FileNotFound)
}