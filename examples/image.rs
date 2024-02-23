use edit_xlsx::{Workbook, WorkbookResult};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    worksheet.insert_image("A1:D4", &"./examples/pics/ferris.png");
    workbook.save_as("examples/image.xlsx")?;
    Ok(())
}