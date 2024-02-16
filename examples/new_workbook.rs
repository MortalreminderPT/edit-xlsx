use edit_xlsx::{Workbook, WorkbookResult};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet()?;
    worksheet.write(1,1,"hello")?;
    workbook.save()?;
    Ok(())
}