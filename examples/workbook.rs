use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor, Write};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    // change the tab ratio
    workbook.set_tab_ratio(50.0)?;
    // recommend to read only
    workbook.read_only_recommended()?;
    let worksheet = workbook.get_worksheet(1)?;
    worksheet.write("A1", "edit excel")?;
    workbook.save()?;
    Ok(())
}