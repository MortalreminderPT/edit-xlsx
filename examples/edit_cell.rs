use edit_xlsx::Workbook;
fn main() {
    let mut workbook = Workbook::from_path("examples/xlsx/edit_cell.xlsx");
    let mut worksheet = workbook.get_worksheet(1).unwrap();
    for row in 1..4 {
        for col in 1..4 {
            worksheet.write(row, col, &format!("writing in ({}, {}) from sheet1", row, col));
        }
    }
    // let mut worksheet = workbook.add_worksheet().unwrap();
    // for row in 1..100 {
    //     for col in 1..100 {
    //         worksheet.write(row, col, &format!("writing in ({}, {}) from sheet2", row, col));
    //     }
    // }
    workbook.save_as("examples/output/edit_cell.xlsx");
}