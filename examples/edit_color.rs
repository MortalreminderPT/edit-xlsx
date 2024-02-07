use edit_xlsx::{Format, Workbook, WorkbookResult, FormatAlign, FormatBorder, FormatColor};
fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::from_path("examples/xlsx/edit_style.xlsx");
    workbook.save_as("examples/output/edit_style.xlsx")?;
    Ok(())
}