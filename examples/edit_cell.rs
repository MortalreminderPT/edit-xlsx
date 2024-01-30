use edit_xlsx::Workbook;
fn main() {
    let mut workbook = Workbook::from_path("examples/edit_cell.xlsx");
    let worksheet = workbook.get_mut_sheet(0).unwrap();
    worksheet.write(1, 1, "after");
    worksheet.write(2, 1, "world");
    worksheet.write(3, 1, "excel");
    worksheet.write(1, 2, "excel");
    worksheet.write(2, 2, "excel");
    worksheet.write(3, 2, "excel");
    worksheet.write(7, 8, "excel");
    worksheet.write(8, 7, "excel");
    worksheet.write(8, 9, "excel");
    workbook.save("examples/edited_cell.xlsx").unwrap();
}
