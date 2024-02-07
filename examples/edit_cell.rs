use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor};
fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::from_path("examples/xlsx/edit_cell.xlsx");
    let mut worksheet = workbook.get_worksheet(1)?;
    // write some text
    worksheet.write(1, 1, "Hello")?;
    worksheet.write(1, 2, "World")?;
    worksheet.write(1, 3, "Rust")?;
    // Adjust font size
    let small = Format::new().set_size(8);
    let medium = Format::new().set_size(16);
    let big = Format::new().set_size(32);
    worksheet.write_with_format(2, 1, "small text", &small)?;
    worksheet.write_with_format(2, 2, "medium text", &medium)?;
    worksheet.write_with_format(2, 3, "big text", &big)?;
    // Change font color
    let red = Format::new().set_color(FormatColor::RGB("00FF7777"));
    let green = Format::new().set_color(FormatColor::RGB("0077FF77"));
    let blue = Format::new().set_color(FormatColor::RGB("007777FF"));
    worksheet.write_with_format(3, 1, "red text", &red)?;
    worksheet.write_with_format(3, 2, "green text", &green)?;
    worksheet.write_with_format(3, 3, "blue text", &blue)?;
    // Change the font style
    let bold = red.set_bold();
    let italic = green.set_italic();
    let underline = blue.set_underline();
    worksheet.write_with_format(4, 1, "red bold text", &bold)?;
    worksheet.write_with_format(4, 2, "green italic text", &italic)?;
    worksheet.write_with_format(4, 3, "blue underline text", &underline)?;
    workbook.save_as("examples/output/edit_cell.xlsx")?;
    Ok(())
}