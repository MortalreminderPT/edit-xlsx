use edit_xlsx::Workbook;

fn main() {
    // Create a new Excel file object.
    let mut workbook = Workbook::from_path("examples/edit_cell.xlsx");

    // Add a worksheet to the workbook.
    let mut worksheet = workbook.add_worksheet();

    // Set the column width for clarity.
    // worksheet.set_column_width(0, 22)?;
    //
    // // Write a string without formatting.
    // worksheet.write(0, 0, "Hello")?;
    workbook.save_as("examples/edited_cell.xlsx");
}