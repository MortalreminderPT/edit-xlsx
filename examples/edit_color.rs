use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::from_path("examples/xlsx/edit.xlsx")?;
    let mut worksheet = workbook.get_worksheet(1)?;
    // Change font color
    let red = Format::default().set_color(FormatColor::RGB("00FF7777"));
    let green = Format::default().set_color(FormatColor::RGB("0077FF77"));
    let blue = Format::default().set_color(FormatColor::RGB("007777FF"));
    worksheet.write_with_format(1, 1, "red text", &red)?;
    worksheet.write_with_format(1, 2, "green text", &green)?;
    worksheet.write_with_format(1, 3, "blue text", &blue)?;
    workbook.save_as("examples/output/edit_color.xlsx")?;
    Ok(())
}