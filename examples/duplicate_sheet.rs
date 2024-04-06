use edit_xlsx::{Workbook, WorkbookResult, Write};

fn main() -> WorkbookResult<()> {
    // from an existed workbook
    let mut workbook = Workbook::from_path("examples/xlsx/accounting.xlsx")?;
    // Use the first worksheet as a template
    let template = workbook.get_worksheet(1)?;
    template.insert_image("I1:L3", &"./examples/pics/ferris.png");
    template.set_name("template")?;

    //
    // Example of using the duplicate_worksheet() function.
    //
    let jan = workbook.duplicate_worksheet(1)?;
    jan.write("A1", "Accounting Journal in Jan.")?;
    jan.set_name("Jan.")?;
    for row in 6..=15 {
        jan.write_row((row, 3), [1, 2, 3, 4, 5, 6, 7, 8, 9, 10].iter())?;
    }
    workbook.save_as("examples/duplicate_sheet.xlsx")?;
    Ok(())
}