use edit_xlsx::{Workbook, WorkbookResult, Read};

fn main() -> WorkbookResult<()> {
    // from an existed workbook
    let mut workbook = Workbook::from_path("examples/xlsx/accounting.xlsx")?;
    workbook.finish();
    // Use the first worksheet as a template
    let sheet1 = workbook.read_worksheet(1)?;
    println!("{}", sheet1.read("B10")?);
    Ok(())
}