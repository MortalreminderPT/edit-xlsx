use edit_xlsx::{FormatColor, Workbook, WorkbookResult};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    worksheet.set_background(&"./examples/pics/rust.png");
    worksheet.write(1, 1, "Sheet1")?;
    workbook.save_as("./examples/output/set_background1.xlsx")?;
    workbook.finish();
    let mut workbook = Workbook::from_path("./examples/output/set_background1.xlsx")?;
    let worksheet = workbook.add_worksheet()?;
    worksheet.write(1, 1, "Sheet2")?;
    worksheet.set_background(&"./examples/pics/ferris.png");
    let tab_color = FormatColor::RGB("00000000");
    // worksheet.set_first_sheet();
    worksheet.set_tab_color(&tab_color);
    workbook.save_as("./examples/output/set_background2.xlsx")?;
    Ok(())
}