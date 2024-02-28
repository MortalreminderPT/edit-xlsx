use edit_xlsx::{Workbook, WorkbookResult, Write};

fn main() -> WorkbookResult<()> {
    // Create a new workbook
    let mut workbook = Workbook::from_path("test/test.xlsx")?;
    let worksheet = workbook.get_worksheet_by_name("应聘表")?;
    worksheet.write("B5", "姓名")?;
    workbook.save_as("test/hello_world.xlsx")?;
    Ok(())
}