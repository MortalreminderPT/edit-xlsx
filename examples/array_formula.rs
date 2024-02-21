use edit_xlsx::{Workbook, WorkbookResult, Write};

fn main() -> WorkbookResult<()> {
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
    worksheet.write_formula("A1", "_xlfn.SUM(B1:C1*B2:C2)")?;
    // Same as above but more verbose.
    worksheet.write_array_formula("A2", "_xlfn.SUM(B1:C1*B2:C2)")?;
    // Write an array formula that returns a range of values
    worksheet.write_array_formula("A5", "_xlfn.TREND(C5:C7,B5:B7)")?;

    workbook.save_as("examples/array_formula.xlsx")?;
    Ok(())
}