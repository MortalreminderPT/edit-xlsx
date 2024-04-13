use edit_xlsx::{Workbook, WorkbookResult};

fn main() -> WorkbookResult<()> {
    // Create a new workbook
    let mut workbook = Workbook::new();
    // Use the default worksheet
    let worksheet = workbook.get_worksheet_mut(1)?;
    // Set the background image.
    // Notice: Only png format is acceptable in the background image
    worksheet.set_background("examples/pics/ferris.png")?;

    workbook.save_as("examples/background.xlsx")?;
    Ok(())
}