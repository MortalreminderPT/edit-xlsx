use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor, Write};
fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::from_path("pic.xlsx")?;
    // let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    // write some text
    for i in (1..50) {
        for j in (1..50) {
            worksheet.write((i, j), format!("{}  {}", i, j))?;
        }
    }
    workbook.save_as("pics.xlsx")?;
    Ok(())
}