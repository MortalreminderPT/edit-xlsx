use edit_xlsx::{Workbook, WorkbookResult};

fn main() -> WorkbookResult<()> {
    // Create a new workbook
    let mut workbook = Workbook::from_path("examples/test/test.xlsx")?;
    let worksheet = workbook.get_worksheet(1)?;
    workbook.save_as("examples/hello_world.xlsx")?;
    Ok(())
}