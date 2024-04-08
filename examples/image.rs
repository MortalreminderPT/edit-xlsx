use edit_xlsx::{Workbook, WorkbookResult};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    // 0.4.0+ Feature: Support for inserting images in the following formats:
    // jpg, png, gif, webp, tif
    worksheet.insert_image("A1:C7", &"./examples/pics/capybara.bmp")?;
    worksheet.insert_image("C7:E14", &"./examples/pics/rust.png")?;
    workbook.save_as("examples/image.xlsx")?;
    Ok(())
}