use edit_xlsx::{Workbook, WorkbookResult};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    worksheet.set_background(&"./examples/pics/rust.png");
    worksheet.write(1, 1, "Sheet1")?;
    workbook.save_as("./examples/output/set_background1.xlsx")?;
    workbook.finish();
    let mut workbook = Workbook::from_path("new(1).xlsx")?;
    let worksheet = workbook.add_worksheet()?;
    worksheet.write(1, 1, "Sheet2")?;
    worksheet.set_background(&"./examples/pics/ferris.png");
    worksheet.set_first_sheet();
    workbook.save_as("./examples/output/set_background2 .xlsx")?;
    Ok(())
}