use edit_xlsx::{Col, Format, Workbook, WorkbookResult, Write};

fn main() -> WorkbookResult<()> {
    // Add the cell formats.
    let format_left_to_right = Format::default()
        .set_reading_order(1);
    let format_right_to_left = Format::default()
        .set_reading_order(2);

    // Create a new workbook
    let mut workbook = Workbook::new();
    // Set up some worksheets and set tab colors
    let worksheet1 = workbook.get_worksheet(1)?;

    // Make the columns wider for clarity.
    worksheet1.set_column("A:A", 25.0)?;

    // Standard direction:         | A1 | B1 | C1 | ...
    worksheet1.write("A1", "نص عربي / English text")?;  // Default direction.
    worksheet1.write_with_format("A2", "نص عربي / English text", &format_left_to_right)?;
    worksheet1.write_with_format("A3", "نص عربي / English text", &format_right_to_left)?;

    let worksheet2 = workbook.add_worksheet()?;
    worksheet2.set_column("A:A", 25.0)?;
    worksheet2.right_to_left();

    // Right to left direction:    ... | C1 | B1 | A1 |
    worksheet2.write("A1", "نص عربي / English text")?;  // Default direction.
    worksheet2.write_with_format("A2", "نص عربي / English text", &format_left_to_right)?;
    worksheet2.write_with_format("A3", "نص عربي / English text", &format_right_to_left)?;

    workbook.save_as("examples/right_to_left.xlsx")?;

    Ok(())
}