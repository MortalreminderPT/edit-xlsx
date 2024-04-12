use edit_xlsx::{Col, Format, FormatColor, Row, Workbook, WorkbookResult, WorkSheet, Write};

fn main() -> WorkbookResult<()> {
    // Create a new workbook
    let mut workbook = Workbook::new();
    // Add a general format
    let bold = Format::default().set_bold();

    //
    // Example 1: A worksheet with outlined rows. It also includes SUBTOTAL()
    // functions so that it looks like the type of automatic outlines that are
    // generated when you use the Excel Data->SubTotals menu item.
    //
    let worksheet1 = workbook.add_worksheet_by_name("Outlined Rows")?;
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
    create_sub_totals(worksheet1)?;

    //
    // Example 2: Create a worksheet with collapsed outlined rows.
    // This is the same as the example 1  except that the all rows are collapsed.
    // Note: We need to indicate the rows that contains the collapsed symbol '+'
    // with the optional parameter, 'collapsed'.
    //
    let worksheet2 = workbook.add_worksheet_by_name("Collapsed Rows 1")?;
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

    // Write the sub-total data that is common to the row examples.
    create_sub_totals(worksheet2)?;


    //
    // Example 3: Create a worksheet with collapsed outlined rows.
    // Same as the example 1  except that the two sub-totals are collapsed.
    //
    let worksheet3 = workbook.add_worksheet_by_name("Collapsed Rows 2")?;
    worksheet3.set_row_level(2, 2)?;
    worksheet3.hide_row(2)?;
    worksheet3.set_row_level(3, 2)?;
    worksheet3.hide_row(3)?;
    worksheet3.set_row_level(4, 2)?;
    worksheet3.hide_row(4)?;
    worksheet3.set_row_level(5, 2)?;
    worksheet3.hide_row(5)?;
    worksheet3.set_row_level(6, 1)?;
    worksheet3.collapse_row(6)?;
    worksheet3.set_row_level(7, 2)?;
    worksheet3.hide_row(7)?;
    worksheet3.set_row_level(8, 2)?;
    worksheet3.hide_row(8)?;
    worksheet3.set_row_level(9, 2)?;
    worksheet3.hide_row(9)?;
    worksheet3.set_row_level(10, 2)?;
    worksheet3.hide_row(10)?;
    worksheet3.set_row_level(11, 1)?;
    worksheet3.collapse_row(11)?;
    // Write the sub-total data that is common to the row examples.
    create_sub_totals(worksheet3)?;


    //
    // Example 4: Create a worksheet with outlined rows.
    // Same as the example 1  except that the two sub-totals are collapsed.
    //
    let worksheet4 = workbook.add_worksheet_by_name("Collapsed Rows 3")?;
    worksheet4.set_row_level(2, 2)?;
    worksheet4.hide_row(2)?;
    worksheet4.set_row_level(3, 2)?;
    worksheet4.hide_row(3)?;
    worksheet4.set_row_level(4, 2)?;
    worksheet4.hide_row(4)?;
    worksheet4.set_row_level(5, 2)?;
    worksheet4.hide_row(5)?;
    worksheet4.set_row_level(6, 1)?;
    worksheet4.hide_row(6)?;
    worksheet4.collapse_row(6)?;
    worksheet4.set_row_level(7, 2)?;
    worksheet4.hide_row(7)?;
    worksheet4.set_row_level(8, 2)?;
    worksheet4.hide_row(8)?;
    worksheet4.set_row_level(9, 2)?;
    worksheet4.hide_row(9)?;
    worksheet4.set_row_level(10, 2)?;
    worksheet4.hide_row(10)?;
    worksheet4.set_row_level(11, 1)?;
    worksheet4.hide_row(11)?;
    worksheet4.collapse_row(11)?;
    worksheet4.collapse_row(12)?;
    // Write the sub-total data that is common to the row examples.
    create_sub_totals(worksheet4)?;



    //
    // Example 5: Create a worksheet with outlined columns.
    //
    let worksheet5 = workbook.add_worksheet_by_name("Outline Columns")?;
    worksheet5.set_row_with_format(1, 15.0, &bold)?;
    worksheet5.write_row("A1", &["Month", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Total"])?;
    worksheet5.set_column_with_format("A:A", 10.0, &bold)?;
    worksheet5.set_column("B:G", 5.0)?;
    worksheet5.set_column_level("B:G", 1)?;
    worksheet5.set_column("H:H", 10.0)?;
    worksheet5.write_column("A2", &["North", "South", "East", "East"])?;
    worksheet5.write_row("B2", &[50, 20, 15, 25, 65, 80])?;
    worksheet5.write_row("B3", &[10, 20, 30, 50, 50, 50])?;
    worksheet5.write_row("B4", &[45, 75, 50, 15, 75, 100])?;
    worksheet5.write_row("B5", &[15, 15, 55, 35, 20, 50])?;
    worksheet5.write_formula("H2", "=SUM(B2:G2)")?;
    worksheet5.write_formula("H3", "=SUM(B3:G3)")?;
    worksheet5.write_formula("H4", "=SUM(B4:G4)")?;
    worksheet5.write_formula("H5", "=SUM(B5:G5)")?;
    worksheet5.write_formula_with_format("H6", "=SUM(H2:H5)", &bold)?;



    //
    // Example 6: Create a worksheet with collapsed outlined columns.
    // This is the same as the previous example except with collapsed columns.
    //
    let worksheet6 = workbook.add_worksheet_by_name("Collapsed Columns")?;
    worksheet6.set_row_with_format(1, 15.0, &bold)?;
    worksheet6.write_row("A1", &["Month", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Total"])?;
    worksheet6.set_column_with_format("A:H", 10.0, &bold)?;
    worksheet6.set_column_level("B:G", 1)?;
    worksheet6.set_column_level("C:F", 2)?;
    worksheet6.set_column_level("D:E", 3)?;
    worksheet6.hide_column("D:E")?;
    worksheet6.collapse_col("D:E")?;
    worksheet6.write_column("A2", &["North", "South", "East", "East"])?;
    worksheet6.write_row("B2", &[50, 20, 15, 25, 65, 80])?;
    worksheet6.write_row("B3", &[10, 20, 30, 50, 50, 50])?;
    worksheet6.write_row("B4", &[45, 75, 50, 15, 75, 100])?;
    worksheet6.write_row("B5", &[15, 15, 55, 35, 20, 50])?;
    worksheet6.write_formula("H2", "=SUM(B2:G2)")?;
    worksheet6.write_formula("H3", "=SUM(B3:G3)")?;
    worksheet6.write_formula("H4", "=SUM(B4:G4)")?;
    worksheet6.write_formula("H5", "=SUM(B5:G5)")?;
    worksheet6.write_formula_with_format("H6", "=SUM(H2:H5)", &bold)?;


    workbook.save_as("examples/outline_collapsed.xlsx")?;
    Ok(())
}

fn create_sub_totals(worksheet: &mut WorkSheet) -> WorkbookResult<()> {
    // Add a general format
    let bold = Format::default().set_bold();

    // Add the data, labels and formulas
    worksheet.write_with_format("A1", "Region", &bold)?;
    worksheet.write("A2", "North")?;
    worksheet.write("A3", "North")?;
    worksheet.write("A4", "North")?;
    worksheet.write("A5", "North")?;
    worksheet.write_with_format("A6", "North Total", &bold)?;

    worksheet.write_with_format("B1", "Sales", &bold)?;
    worksheet.write("B2", 1000)?;
    worksheet.write("B3", 1200)?;
    worksheet.write("B4", 900)?;
    worksheet.write("B5", 1200)?;
    worksheet.write_formula_with_format("B6", "=SUBTOTAL(9,B2:B5)", &bold)?;
    worksheet.write("A7", "South")?;
    worksheet.write("A8", "South")?;
    worksheet.write("A9", "South")?;
    worksheet.write("A10", "South")?;
    worksheet.write_with_format("A11", "South Total", &bold)?;
    worksheet.write("B7", 400)?;
    worksheet.write("B8", 600)?;
    worksheet.write("B9", 500)?;
    worksheet.write("B10", 600)?;
    worksheet.write_formula_with_format("B11", "=SUBTOTAL(9,B7:B10)", &bold)?;
    worksheet.write_with_format("A12", "Grand Total", &bold)?;
    worksheet.write_formula_with_format("B12", "=SUBTOTAL(9,B2:B10)", &bold)?;
    Ok(())
}