use edit_xlsx::{Workbook, WorkbookResult, Col, Read, Write, Row};

fn main() -> WorkbookResult<()> {
    // from an existed workbook
    let reading_book = Workbook::from_path("./examples/xlsx/accounting.xlsx")?;
    // Read the first sheet
    let reading_sheet = reading_book.get_worksheet(1)?;
    let mut writing_book = Workbook::new();
    let writing_sheet = writing_book.get_worksheet_mut(1)?;
    writing_sheet.set_default_row(writing_sheet.get_default_row());
    // let bg_format = reading_sheet.read_format();

    // Synchronous column width
    let mut columns_map = reading_sheet.get_columns_with_format("A:XFD")?;
    println!("{:#?}", columns_map);
    columns_map.iter_mut().for_each(|(col_range, (column, format))| {
        if let Some(format) = format {
            writing_sheet.set_columns_with_format(col_range, column, format).unwrap()
        } else {
            writing_sheet.set_columns(col_range, column).unwrap()
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
    writing_book.save_as("./examples/read.xlsx")?;
    Ok(())
}