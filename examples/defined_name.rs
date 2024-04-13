use edit_xlsx::{Col, Format, FormatColor, Workbook, WorkbookResult, WorkSheet, WorkSheetResult, Write};

fn main() -> WorkbookResult<()> {
    // Create a new workbook
    let mut workbook = Workbook::new();
    workbook.add_worksheet()?;

    // Define some global/workbook names.
    workbook.define_name("Exchange_rate", "0.96")?;
    workbook.define_name("Sales", "Sheet1!$G$1:$H$10")?;
    // Define a local/worksheet name. Over-rides the "Sales" name above.
    workbook.define_local_name("Sales", "Sheet2!$G$1:$G$10", 2)?;

    // Write some text in the file and one of the defined names in a formula.
    let sales = ["Apple", "Grape", "Pear", "Banana", "Apple", "Grape", "Pear", "Banana", "Banana", "Pear"];
    let units = [10, 12, 32, 16, 13, 50, 25, 8, 33, 95];
    for worksheet in workbook.worksheets_mut() {
        worksheet.set_columns_width("A:B", 40.0)?;
        worksheet.set_columns_width("F:F", 40.0)?;
        worksheet.write("A1", "This worksheet contains some defined names.")?;
        worksheet.write("B1", "Show defined name Sales on the right->")?;
        worksheet.write_formula("C1", "=Sales")?;
        // In older versions of Excel, you could use the write_old_formula method:
        // worksheet.write_old_formula("C1", "=Sales")?;
        worksheet.write("A2", "See Formulas -> Name Manager above.")?;
        worksheet.write("A3", "Example formula in cell B3 ->")?;
        worksheet.write_formula("B3", "=Exchange_rate")?;
        // In older versions of Excel, you could use the write_old_formula method:
        // worksheet.write_old_formula("B3", "=Exchange_rate")?;
        worksheet.write("F1", "Fill in some arrays on the right->")?;
        worksheet.write_column("G1", &sales)?;
        worksheet.write_column("H1", &units)?;
    }

    workbook.save_as("examples/defined_name.xlsx")?;
    Ok(())
}
