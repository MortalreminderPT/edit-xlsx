use std::collections::HashMap;
use edit_xlsx::{Workbook, WorkbookResult, Col, Write};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;

    // Write strings that looks like numbers. This will cause an Excel warning.
    worksheet.write_string("C2", "123".to_string())?;
    worksheet.write_string("C3", "123".to_string())?;

    // Write a divide by zero formula. This will also cause an Excel warning.
    worksheet.write_formula("C5", "=1/0")?;
    worksheet.write_formula("C6", "=1/0")?;
    // In older versions of Excel, you could use the write_old_formula method:
    // worksheet.write_old_formula("C5", "=1/0")?;
    // worksheet.write_old_formula("C6", "=1/0")?;

    // Turn off some of the warnings:
    let mut error_map = HashMap::new();
    error_map.insert("number_stored_as_text", "C3");
    error_map.insert("eval_error", "C6");
    worksheet.ignore_errors(error_map);

    // Write some descriptions for the cells and make the column wider for clarity.
    worksheet.set_column("B:B", 16.0)?;
    worksheet.write("B2", "Warning:")?;
    worksheet.write("B3", "Warning turned off:")?;
    worksheet.write("B5", "Warning:")?;
    worksheet.write("B6", "Warning turned off:")?;

    workbook.save_as("examples/ignore_errors.xlsx")?;
    Ok(())
}