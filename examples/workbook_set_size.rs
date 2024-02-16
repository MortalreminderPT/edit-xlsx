use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    workbook.set_size(1200, 800)?;
    workbook.save()?;
    Ok(())
}