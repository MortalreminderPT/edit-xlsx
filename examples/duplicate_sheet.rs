use edit_xlsx::{Format, Workbook, WorkbookResult, WorkSheet, WorkSheetResult, Write};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::from_path("Z:\\accounting.xlsx")?;
    let worksheet = workbook.get_worksheet(1);
    workbook.save_as("Z:\\accounting2.xlsx")?;
    Ok(())
}