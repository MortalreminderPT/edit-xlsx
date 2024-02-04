use edit_xlsx::Workbook;

fn main() {
    // Create a new Excel file object.
    let mut workbook = Workbook::from_path("examples/xlsx/edit_sheet.xlsx");
    // Add a worksheet to the workbook.
    let mut worksheet = workbook.add_worksheet();
    for i in 1..1000 {
        let mut worksheet = workbook.add_worksheet().unwrap();
        worksheet.write(1, 1, &format!("sheet{}", i));
    }
    workbook.save_as("examples/output/edit_sheet.xlsx");
}