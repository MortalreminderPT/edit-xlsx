use edit_xlsx::{Col, Format, Row, Workbook, WorkbookResult, Write};

fn main() -> WorkbookResult<()> {
    // Create a new workbook
    let mut workbook = Workbook::new();
    let worksheet1 = workbook.add_worksheet_by_name("Outlined Rows")?;
    // Add a general format
    let bold = Format::default().set_bold();

    //
    // Example 1: A worksheet with outlined rows. It also includes SUBTOTAL()?;    // functions so that it looks like the type of automatic outlines that are
    // generated when you use the Excel Data->SubTotals menu item.
    //
    // For outlines the important parameters are 'level' and 'hidden'. Rows with
    // the same 'level' are grouped together. The group will be collapsed if
    // 'hidden' is enabled. The parameters 'height' and 'cell_format' are assigned
    // default values if they are None.
    //
    worksheet1.set_row_level(2, 2)?;
    worksheet1.set_row_level(3, 2)?;
    worksheet1.set_row_level(4, 2)?;
    worksheet1.set_row_level(5, 2)?;
    worksheet1.set_row_level(6, 1)?;
    worksheet1.set_row_level(7, 2)?;
    worksheet1.set_row_level(8, 2)?;
    worksheet1.set_row_level(9, 2)?;
    worksheet1.set_row_level(10, 2)?;
    worksheet1.set_row_level(11, 1)?;

    // Adjust the column width for clarity
    worksheet1.set_columns_width("A:A", 20.0)?;

    // Add the data, labels and formulas
    worksheet1.write_with_format("A1", "Region", &bold)?;
    worksheet1.write("A2", "North")?;
    worksheet1.write("A3", "North")?;
    worksheet1.write("A4", "North")?;
    worksheet1.write("A5", "North")?;
    worksheet1.write_with_format("A6", "North Total", &bold)?;

    worksheet1.write_with_format("B1", "Sales", &bold)?;
    worksheet1.write("B2", 1000)?;
    worksheet1.write("B3", 1200)?;
    worksheet1.write("B4", 900)?;
    worksheet1.write("B5", 1200)?;
    worksheet1.write_formula_with_format("B6", "=SUBTOTAL(9,B2:B5)", &bold)?;
    worksheet1.write("A7", "South")?;
    worksheet1.write("A8", "South")?;
    worksheet1.write("A9", "South")?;
    worksheet1.write("A10", "South")?;
    worksheet1.write_with_format("A11", "South Total", &bold)?;
    worksheet1.write("B7", 400)?;
    worksheet1.write("B8", 600)?;
    worksheet1.write("B9", 500)?;
    worksheet1.write("B10", 600)?;
    worksheet1.write_formula_with_format("B11", "=SUBTOTAL(9,B7:B10)", &bold)?;
    worksheet1.write_with_format("A12", "Grand Total", &bold)?;
    worksheet1.write_formula_with_format("B12", "=SUBTOTAL(9,B2:B10)", &bold)?;


    //
    // Example 2: A worksheet with outlined rows. This is the same as the
    // previous example except that the rows are collapsed.
    // Note: We need to indicate the rows that contains the collapsed symbol '+'
    // with the optional parameter, 'collapsed'. The group will be then be
    // collapsed if 'hidden' is True.
    //
    let worksheet2 = workbook.add_worksheet_by_name("Collapsed Rows")?;
    worksheet2.set_row_level(2, 2)?;
    worksheet2.hide_row(2)?;
    worksheet2.set_row_level(3, 2)?;
    worksheet2.hide_row(3)?;
    worksheet2.set_row_level(4, 2)?;
    worksheet2.hide_row(4)?;
    worksheet2.set_row_level(5, 2)?;
    worksheet2.hide_row(5)?;
    worksheet2.set_row_level(6, 1)?;
    worksheet2.hide_row(6)?;
    worksheet2.set_row_level(7, 2)?;
    worksheet2.hide_row(7)?;
    worksheet2.set_row_level(8, 2)?;
    worksheet2.hide_row(8)?;
    worksheet2.set_row_level(9, 2)?;
    worksheet2.hide_row(9)?;
    worksheet2.set_row_level(10, 2)?;
    worksheet2.hide_row(10)?;
    worksheet2.set_row_level(11, 1)?;
    worksheet2.hide_row(11)?;
    worksheet2.collapse_row(12)?;

    // Adjust the column width for clarity
    worksheet2.set_columns_width("A:A", 20.0)?;

    // Add the data, labels and formulas
    worksheet2.write_with_format("A1", "Region", &bold)?;
    worksheet2.write("A2", "North")?;
    worksheet2.write("A3", "North")?;
    worksheet2.write("A4", "North")?;
    worksheet2.write("A5", "North")?;
    worksheet2.write_with_format("A6", "North Total", &bold)?;

    worksheet2.write_with_format("B1", "Sales", &bold)?;
    worksheet2.write("B2", 1000)?;
    worksheet2.write("B3", 1200)?;
    worksheet2.write("B4", 900)?;
    worksheet2.write("B5", 1200)?;
    worksheet2.write_formula_with_format("B6", "=SUBTOTAL(9,B2:B5)", &bold)?;
    worksheet2.write("A7", "South")?;
    worksheet2.write("A8", "South")?;
    worksheet2.write("A9", "South")?;
    worksheet2.write("A10", "South")?;
    worksheet2.write_with_format("A11", "South Total", &bold)?;
    worksheet2.write("B7", 400)?;
    worksheet2.write("B8", 600)?;
    worksheet2.write("B9", 500)?;
    worksheet2.write("B10", 600)?;
    worksheet2.write_formula_with_format("B11", "=SUBTOTAL(9,B7:B10)", &bold)?;
    worksheet2.write_with_format("A12", "Grand Total", &bold)?;
    worksheet2.write_formula_with_format("B12", "=SUBTOTAL(9,B2:B10)", &bold)?;

    //
    // Example 3: Create a worksheet with outlined columns.
    //
    let worksheet3 = workbook.add_worksheet_by_name("Outline Columns")?;
    // Add bold format to the first row.
    worksheet3.set_row_with_format(1, 15.0, &bold)?;
    worksheet3.write_row("A1", &["Month", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Total"])?;
    worksheet3.set_columns_width_with_format("A:A", 10.0, &bold)?;
    worksheet3.set_columns_width("B:G", 10.0)?;
    worksheet3.set_columns_level("B:G", 1)?;
    worksheet3.set_columns_width("H:H", 10.0)?;
    worksheet3.write_column("A2", &["North", "South", "East", "East"])?;
    worksheet3.write_row("B2", &[50, 20, 15, 25, 65, 80])?;
    worksheet3.write_row("B3", &[10, 20, 30, 50, 50, 50])?;
    worksheet3.write_row("B4", &[45, 75, 50, 15, 75, 100])?;
    worksheet3.write_row("B5", &[15, 15, 55, 35, 20, 50])?;
    worksheet3.write_formula("H2", "=SUM(B2:G2)")?;
    worksheet3.write_formula("H3", "=SUM(B3:G3)")?;
    worksheet3.write_formula("H4", "=SUM(B4:G4)")?;
    worksheet3.write_formula("H5", "=SUM(B5:G5)")?;
    worksheet3.write_formula_with_format("H6", "=SUM(H2:H5)", &bold)?;

    //
    // Example 4: Show all possible outline levels.
    //
    let levels = [
        "Level 1",
        "Level 2",
        "Level 3",
        "Level 4",
        "Level 5",
        "Level 6",
        "Level 7",
        "Level 6",
        "Level 5",
        "Level 4",
        "Level 3",
        "Level 2",
        "Level 1",
    ];
    let worksheet4 = workbook.add_worksheet_by_name("Outline levels")?;
    worksheet4.write_column("A1", &levels)?;

    worksheet4.set_row_level(1, 1)?;
    worksheet4.set_row_level(2, 2)?;
    worksheet4.set_row_level(3, 3)?;
    worksheet4.set_row_level(4, 4)?;
    worksheet4.set_row_level(5, 5)?;
    worksheet4.set_row_level(6, 6)?;
    worksheet4.set_row_level(7, 7)?;
    worksheet4.set_row_level(8, 6)?;
    worksheet4.set_row_level(9, 5)?;
    worksheet4.set_row_level(10, 4)?;
    worksheet4.set_row_level(11, 3)?;
    worksheet4.set_row_level(12, 2)?;
    worksheet4.set_row_level(13, 1)?;

    workbook.save_as("examples/outline.xlsx")?;
    Ok(())
}