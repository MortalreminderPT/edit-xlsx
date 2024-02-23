use edit_xlsx::{Workbook, WorkbookResult, Write};

fn main() -> WorkbookResult<()> {
    // Create a new workbook
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    worksheet.write("A1", "Hello world")?;
    workbook.save_as("examples/hello_world.xlsx")?;
    Ok(())
}
