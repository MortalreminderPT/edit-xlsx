use edit_xlsx::{Col, Format, FormatColor, Workbook, WorkbookResult, WorkSheet, WorkSheetResult, Write};

fn main() -> WorkbookResult<()> {
    // Create a new workbook
    let mut workbook = Workbook::from_path("dynamic_arrays.xlsx")?;
    workbook.save_as("test/test.xlsx")?;
    Ok(())
}
