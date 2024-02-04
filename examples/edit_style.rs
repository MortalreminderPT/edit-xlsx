use edit_xlsx::Workbook;

fn main() {
    let mut workbook = Workbook::from_path("examples/xlsx/edit_style.xlsx");
    let mut worksheet = workbook.get_worksheet(1).unwrap();

    workbook.save_as("examples/output/edit_style.xlsx");
}