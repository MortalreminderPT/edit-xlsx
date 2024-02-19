use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor, Write};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    // write string
    worksheet.write_v2((1, 1), "aa")?;
    worksheet.write_v2("A3", "bb")?;
    worksheet.write_v2("4B", "cc")?;
    workbook.save()?;
    Ok(())
}