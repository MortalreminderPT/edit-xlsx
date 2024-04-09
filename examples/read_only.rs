use edit_xlsx::{Workbook, WorkbookResult, Read, Write, Col};

fn main() -> WorkbookResult<()> {
    // from an existed workbook
    let mut reading_book = Workbook::from_path("examples/xlsx/accounting.xlsx")?;
    reading_book.finish();
    // Read the first sheet
    let reading_sheet = reading_book.read_worksheet(1)?;
    let mut writing_book = Workbook::new();
    let writing_sheet = writing_book.get_worksheet(1)?;
    for row in 1..=reading_sheet.max_row() {
        for col in 1..=reading_sheet.max_column() {
            let text = reading_sheet.read((row, col)).unwrap_or_default();
            let format = reading_sheet.read_format((row, col)).unwrap_or_default();
            writing_sheet.write_with_format((row, col), text, &format).unwrap()
        }
    }
    writing_book.save_as("./examples/new.xlsx")?;
    Ok(())
}