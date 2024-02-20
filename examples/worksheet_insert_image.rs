use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor, Row};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    worksheet.insert_image((1, 1), &"./examples/pics/ferris.png");
    workbook.save()?;
    Ok(())
}