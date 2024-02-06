use edit_xlsx::{Workbook, WorkbookResult};

fn main() -> WorkbookResult<()> {
    // Create a new Excel file object.
    let mut workbook = Workbook::from_path("examples/xlsx/edit_sheet.xlsx");
    // Add worksheets to the workbook.
    for i in 1..10 {
        let mut worksheet = workbook.add_worksheet().unwrap();
        worksheet.write(1, 1, &format!("Text in Sheet{}", worksheet.id()))?;
    }
    workbook.save_as("examples/output/edit_sheet.xlsx")?;
    Ok(())
}