use edit_xlsx::{Workbook, WorkbookResult, Col, Read, Write, Row, WorkSheetResult};

fn main() -> WorkbookResult<()> {
    // from an existed workbook
    let mut reading_book = Workbook::from_path("examples/new.xlsx")?;
    reading_book.finish();
    // Read the first sheet
    let reading_sheet = reading_book.read_worksheet(1)?;
    let mut writing_book = Workbook::new();
    let writing_sheet = writing_book.get_worksheet(1)?;
    writing_sheet.set_default_row(writing_sheet.get_default_row());
    // let bg_format = reading_sheet.read_format();

    // Synchronous column width
    let mut widths = reading_sheet.get_columns_with_format((1, 1, 1, 16384))?;
    widths.iter_mut().for_each(|(min, max, column, format)| {
        if let Some(format) = format {
            writing_sheet.set_columns_with_format((1, *min, 1, *max), column, format).unwrap()
        } else {
            writing_sheet.set_columns((1, *min, 1, *max), column).unwrap()
        }
    });

    // Read then write text and format
    for row in 1..=reading_sheet.max_row() {
        for col in 1..=reading_sheet.max_column() {
            match (reading_sheet.read((row, col)), reading_sheet.read_format((row, col))) {
                (Ok(text), Ok(format)) => {
                    writing_sheet.write_with_format((row, col), text, &format).unwrap();
                }
                (Ok(text), _) => {
                    writing_sheet.write((row, col), text).unwrap();
                }
                (_, Ok(format)) => {
                    writing_sheet.write_with_format((row, col), "", &format).unwrap();
                }
                _ => {}
            }
            if let Ok(height) = writing_sheet.get_row(row) {
                writing_sheet.set_row(row, height)?;
            }
        }
    }
    writing_book.save_as("./examples/new2.xlsx")?;
    Ok(())
}