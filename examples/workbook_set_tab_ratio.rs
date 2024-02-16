use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    workbook.set_tab_ratio(50.0)?;
    workbook.save()?;
    Ok(())
}