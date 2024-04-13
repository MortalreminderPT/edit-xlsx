use edit_xlsx::{Workbook, WorkbookResult, WorkSheetCol, Read, Write, WorkSheetRow};

/// NOTICE: Reading is an experimental feature,
/// in the current version it performs better for reading and overwriting text, styles,
/// and worse for formulas, links, etc.

fn main() -> WorkbookResult<()> {
    // Read an existed workbook
    let reading_book = Workbook::from_path("./tests/xlsx/accounting.xlsx")?;
    let reading_sheet = reading_book.get_worksheet_by_name("worksheet")?;
    // Create a new workbook to write
    let mut writing_book = Workbook::new();
    let writing_sheet = writing_book.get_worksheet_mut(1)?;

    // Synchronous column width and format
    let columns_map = reading_sheet.get_columns_with_format("A:XFD")?;
    columns_map.iter().for_each(|(col_range, (column, format))| {
        if let Some(format) = format {
            // if col format exists, write it to writing_sheet
            writing_sheet.set_columns_with_format(col_range, column, format).unwrap()
        } else {
            writing_sheet.set_columns(col_range, column).unwrap()
        }
    });

    // Synchronous row height and format
    writing_sheet.set_default_row(writing_sheet.get_default_row());
    for row_number in 1..=reading_sheet.max_row() {
        let (row, format) = reading_sheet.get_row_with_format(row_number)?;
        if let Some(format) = format {
            // if col format exists, write it to writing_sheet
            writing_sheet.set_row_with_format(row_number, &row, &format)?;
        } else {
            writing_sheet.set_row(row_number, &row)?;
        }
    }

    // Read then write text and format
    for row in 1..=reading_sheet.max_row() {
        for col in 1..=reading_sheet.max_column() {
            if let Ok(cell) = reading_sheet.read((row, col)) {
                writing_sheet.write_cell((row, col), &cell)?;
            }
        }
    }

    writing_book.save_as("./examples/read_and_copy.xlsx")?;
    Ok(())
}