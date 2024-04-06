use edit_xlsx::{Workbook, WorkbookResult, Write};

fn main() -> WorkbookResult<()> {
    // from an existed workbook
    let mut workbook = Workbook::from_path("examples/xlsx/accounting.xlsx")?;
    // Use the first worksheet as a template
    let template = workbook.get_worksheet(1)?;
    template.insert_image("I1:L3", &"./examples/pics/ferris.png");
    template.set_name("template")?;
    // Deselect and hide the template.
    template.deselect();
    template.hide();
    //
    // Example of using the duplicate_worksheet() function.
    //
    let jan = workbook.duplicate_worksheet(1)?;
    jan.write("A1", "Accounting Journal in Jan.")?;
    jan.set_name("Jan.")?;
    for row in 6..=15 {
        jan.write_row((row, 3), [1, 2, 3, 4, 5, 6, 7, 8, 9, 10].iter())?;
    }
    let feb = workbook.duplicate_worksheet(1)?;
    feb.write("A1", "Accounting Journal in Feb.")?;
    feb.set_name("Feb.")?;
    for row in 6..=15 {
        feb.write_row((row, 3), [2, 4, 6, 8, 10, 12, 14, 16, 18, 20].iter())?;
    }
    // activate the Feb. sheet.
    feb.activate();
    let mar = workbook.duplicate_worksheet(1)?;
    mar.write("A1", "Accounting Journal in Mar.")?;
    mar.set_name("Mar.")?;
    for col in 'C'..='L' {
        mar.write_column(&format!("{col}6"), [1, 2, 3, 4, 5, 6, 7, 8, 9, 10].iter())?;
    }
    workbook.save_as("examples/duplicate_sheet.xlsx")?;
    Ok(())
}