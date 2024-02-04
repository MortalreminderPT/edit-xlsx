use edit_xlsx::{Format, Workbook};

fn main() {
    let mut workbook = Workbook::from_path("examples/xlsx/edit_style.xlsx");
    let mut worksheet = workbook.get_worksheet(1).unwrap();
    for row in 1..9 {
        for col in 1..9 {
            worksheet.write_with_format(
                row,
                col,
                &format!("writing in ({}, {}) from sheet1", row, col),
                match row % 3 {
                    0 => Format::new().set_bold(),
                    1 => Format::new().set_bold().set_underline(),
                    2.. => Format::new().set_bold().set_italic().set_underline(),
                }
            );
        }
    }
    workbook.save_as("examples/output/edit_style.xlsx");
}