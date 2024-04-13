use edit_xlsx::{Workbook, WorkbookResult, Write};

fn main() -> WorkbookResult<()> {
    // Create a new workbook
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet_mut(1)?;


    // Write a Lambda function to convert Fahrenheit to Celsius to a cell.
    //
    // Note that the lambda function parameters must be prefixed with
    // "_xlpm.". These prefixes won't show up in Excel.
    worksheet.write_formula("A1", "_xlfn.LAMBDA(_xlpm.a, _xlpm.b, SQRT((_xlpm.a^2+_xlpm.b^2)))(3, 4)")?;


    // The user defined name needs to be written explicitly as a dynamic array
    // formula.
    let a = 6;
    let b = 8;
    worksheet.write_formula("A2", &format!("=HYPOTENUSE({a}, {b})"))?;

    // Create the same formula (without an argument) as a defined name and use that
    // to calculate a value.
    //
    // Note that the formula name is prefixed with "_xlfn." (this is normally
    // converted automatically by write_formula() but isn't for defined names)
    // and note that the lambda function parameters are prefixed with
    // "_xlpm.". These prefixes won't show up in Excel.
    workbook.define_name("HYPOTENUSE", "_xlfn.LAMBDA(_xlpm.a, _xlpm.b, SQRT((_xlpm.a^2+_xlpm.b^2)))")?;


    workbook.save_as("examples/lambda.xlsx")?;
    Ok(())
}