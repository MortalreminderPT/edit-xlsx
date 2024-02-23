use edit_xlsx::{Col, Row, Workbook, WorkbookResult, Write};

fn main() -> WorkbookResult<()> {
    // Create a new workbook
    let mut workbook = Workbook::new();

    let worksheet = workbook.get_worksheet(1)?;

    // Write some data.
    worksheet.write("D1", "Some hidden columns.")?;
    worksheet.write("A8", "Some hidden rows.")?;

    // Hide all rows without data.
    worksheet.hide_unused_rows(true);

    // Set the height of empty rows that we do want to display even if it is
    // the default height.
    for row in 2..=7 {
        worksheet.set_row(row, 15.0)?;
    }

    // Columns can be hidden explicitly. This doesn't increase the file size..
    worksheet.hide_column("G:XFD")?;

    // Note: hide_column will overwrite the column settings and will be overwritten by set_column
    worksheet.set_column("G:XFD", 1.0)?;

    workbook.save_as("examples/hide_row_col.xlsx")?;
    Ok(())
}