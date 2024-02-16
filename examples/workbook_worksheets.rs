use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    workbook.add_worksheet()?;
    workbook.add_worksheet_by_name("Foglio2")?;
    workbook.add_worksheet_by_name("Data")?;
    workbook.add_worksheet()?;

    for worksheet in workbook.worksheets() {
        worksheet.write(1, 1, "hello")?;
    }
    workbook.save()?;
    Ok(())
}