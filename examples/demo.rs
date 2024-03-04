use edit_xlsx::{Workbook, WorkbookResult, Write};

fn main() -> WorkbookResult<()> {
    // Create a new workbook
    let mut workbook = Workbook::from_path("test/test2.xlsx")?;
    let worksheet = workbook.get_worksheet(4)?;
    let worksheet = workbook.add_worksheet()?;
    worksheet.write("B5", "职位名称")?;
    workbook.save_as("test/hello_world.xlsx")?;
    // workbook.save_as("test/hello_world.xlsx")?;
    // workbook.save_as("test/hello_world.xlsx")?;
    Ok(())
}