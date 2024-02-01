use edit_xlsx::Workbook;
fn main(){
    let mut workbook = Workbook::from_path("examples/edit_cell.xlsx");
    // let mut worksheet = workbook.get_sheet_by_name_mut("Sheet1");
    let worksheet = workbook.get_sheet_mut(1).unwrap();
    worksheet.write(1, 1, "after");
    worksheet.write(2, 1, "world");
    worksheet.write(3, 1, "excel");
    worksheet.write(1, 2, "excel");
    worksheet.write(2, 2, "excel");
    worksheet.write(3, 2, "excel");
    worksheet.write(7, 8, "excel");
    worksheet.write(8, 7, "excel");
    worksheet.write(8, 9, "excel");

    // let mut worksheet = workbook.get_sheet_mut(2);
    for row in (1..100).rev() {
        for col in (1..100).rev() {
            worksheet.write(row, col, &format!("writing in ({}, {})", row, col));
        }
    }
    workbook.save("examples/edited_cell.xlsx");
}