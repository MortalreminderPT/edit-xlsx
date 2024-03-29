use edit_xlsx::{Workbook, WorkbookResult};

fn main() -> WorkbookResult<()> {
    // Create a new workbook
    let mut workbook = Workbook::new();
    // Use the default worksheet
    let worksheet = workbook.get_worksheet(1)?;
    // Set the background image.
    worksheet.set_background("examples/pics/ferris.png")?;

    workbook.save_as("examples/background.xlsx")?;
    Ok(())
}