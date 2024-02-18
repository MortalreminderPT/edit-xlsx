use edit_xlsx::{Workbook, WorkbookResult};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::from_path("new2.xlsx")?;
    // let mut workbook = Workbook::new();
    // let worksheet = workbook.get_worksheet(1)?;
    // worksheet.set_background(&"./examples/pics/rust.png");
    let worksheet = workbook.add_worksheet()?;
    worksheet.set_background(&"./examples/pics/rust.jpg");
    let worksheet = workbook.add_worksheet()?;
    worksheet.set_background(&"examples/pics/ferris.png");
    let worksheet = workbook.add_worksheet()?;
    worksheet.set_background(&"./examples/pics/rust.jpg");
    let worksheet = workbook.add_worksheet()?;
    worksheet.set_background(&"./examples/pics/ferris.png");
    // let worksheet = workbook.add_worksheet()?;
    // worksheet.set_background(&"./examples/pics/rust.png");
    // worksheet.write(1, 1, 1)?;
    workbook.save_as("new3.xlsx")?;
    Ok(())
}