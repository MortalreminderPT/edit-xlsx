use edit_xlsx::{Format, FormatBorder, Workbook};

fn main() {
    let mut workbook = Workbook::from_path("examples/xlsx/edit_style.xlsx");
    let mut worksheet = workbook.get_worksheet(1).unwrap();
    let bold_format = Format::new().set_bold().set_underline().set_italic();

    worksheet.write_with_format(1, 1, "An example of text font style", &bold_format);
    let border_format = Format::new().set_border(FormatBorder::MediumDashed);
    worksheet.write_with_format(1, 2, "An example of border style", &border_format);
    let underline_italic_format = Format::new().set_underline().set_italic();
    let bold_underline_format = Format::new().set_bold().set_underline();
    for row in 3..20 {
        for col in 3..20 {
            worksheet.write_with_format(
                row,
                col,
                &format!("writing in ({}, {}) from sheet1", row, col),
                match row % 3 {
                    0 => &bold_format,
                    1 => &bold_format,
                    2.. => &bold_underline_format,
                }
            );
        }
    }
    workbook.save_as("examples/output/edit_style.xlsx");
}