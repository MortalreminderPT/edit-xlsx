use edit_xlsx::{Workbook, WorkbookResult, Read};

fn main() -> WorkbookResult<()> {
    // from an existed workbook
    let mut workbook = Workbook::from_path("examples/xlsx/accounting.xlsx")?;
    workbook.finish();
    // Read the first sheet
    let sheet = workbook.read_worksheet(1)?;
    for row in 1..=sheet.max_row() {
        for col in 1..=sheet.max_column() {
            print!("{}\t", sheet.read((row, col)).unwrap_or_default());
        }
        println!("\n")
    }
    Ok(())
}