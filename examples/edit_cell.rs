use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor, Write};
fn main() -> WorkbookResult<()> {
    // let mut workbook = Workbook::from_path("examples/xlsx/edit.xlsx")?;
    let mut workbook = Workbook::new();
    let mut worksheet = workbook.get_worksheet(1)?;
    // write some text
    for i in (1..50).rev() {
        for j in (1..50).rev() {
            worksheet.write((i, j), format!("{}  {}", worksheet.max_row(), worksheet.max_column()))?;
        }
    }
    workbook.save_as("examples/output/edit_cell.xlsx")?;
    Ok(())
}