use edit_xlsx::{Format, Workbook, WorkbookResult, FormatAlign, FormatBorder, FormatColor};
fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::from_path("examples/xlsx/edit_style.xlsx");
    let mut worksheet = workbook.get_worksheet(1)?;
    // adjust the text align
    let center = Format::new().set_align(FormatAlign::VerticalCenter).set_align(FormatAlign::Center);
    let left_top = Format::new().set_align(FormatAlign::Left).set_align(FormatAlign::Top);
    let right_bottom = Format::new().set_align(FormatAlign::Right).set_align(FormatAlign::Bottom);
    worksheet.write_with_format(1, 1, "center", &center)?;
    worksheet.write_with_format(1, 2, "left top", &left_top)?;
    worksheet.write_with_format(1, 3, "right bottom", &right_bottom)?;
    // add borders
    let thin_border = Format::new().set_border(FormatBorder::Double);
    worksheet.write_with_format(2, 1, "bordered text", &thin_border)?;
    let thin_border = Format::new().set_border_bottom(FormatBorder::Double);
    worksheet.write_with_format(2, 2, "bordered text", &thin_border)?;
    let thin_border = Format::new().set_border_right(FormatBorder::Double);
    worksheet.write_with_format(2, 3, "bordered text", &thin_border)?;
    // add background
    let red_background = Format::new().set_background_color(FormatColor::RGB("00FF7777"));
    worksheet.write_with_format(3, 1, "red", &red_background)?;
    workbook.save_as("examples/output/edit_style.xlsx")?;
    Ok(())
}