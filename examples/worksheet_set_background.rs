use edit_xlsx::{Workbook, WorkbookResult};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    worksheet.set_background("hello");
    worksheet.write(1,1,1)?;
    workbook.save()?;
    Ok(())
}