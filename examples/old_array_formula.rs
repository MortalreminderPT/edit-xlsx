use edit_xlsx::{Workbook, WorkbookResult, Write};

fn main() -> WorkbookResult<()> {
    /// In Office 365, many new functions have emerged.
    /// However, these functions may conflict with older versions of Excel.
    /// You can use the 'write_old_formula' method,
    /// which will minimize potential conflicts as much as possible.

    // Create a new workbook
    let mut workbook = Workbook::new();
    // Use the default worksheet
    let worksheet = workbook.get_worksheet(1)?;
    // Write some test data.
    worksheet.write("B1", 500)?;
    worksheet.write("B2", 10)?;
    worksheet.write("B5", 1)?;
    worksheet.write("B6", 2)?;
    worksheet.write("B7", 3)?;
    worksheet.write("C1", 300)?;
    worksheet.write("C2", 15)?;
    worksheet.write("C5", 20234)?;
    worksheet.write("C6", 21003)?;
    worksheet.write("C7", 10000)?;
    // Write an array formula that returns a single value
    worksheet.write_old_formula("A1", "_xlfn.SUM(B1:C1*B2:C2)")?;
    // Same as above but more verbose.
    worksheet.write_old_formula("A2", "_xlfn.SUM(B1:C1*B2:C2)")?;
    // Write an array formula that returns a range of values
    worksheet.write_old_formula("A5", "_xlfn.TREND(C5:C7,B5:B7)")?;

    workbook.save_as("examples/old_array_formula.xlsx")?;
    Ok(())
}