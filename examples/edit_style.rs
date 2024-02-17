use edit_xlsx::{Format, Workbook, WorkbookResult, FormatAlignType, FormatBorderType, FormatColor};


fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::from_path("examples/xlsx/new.xlsx")?;
    let worksheet = workbook.get_worksheet(1)?;
    // adjust the text align
    let center = Format::default().set_align(FormatAlignType::VerticalCenter).set_align(FormatAlignType::Center);
    let left_top = Format::default().set_align(FormatAlignType::Left).set_align(FormatAlignType::Top);
    let right_bottom = Format::default().set_align(FormatAlignType::Right).set_align(FormatAlignType::Bottom);
    worksheet.write_with_format(1, 1, "center", &center)?;
    worksheet.write_with_format(1, 2, "left top", &left_top)?;
    worksheet.write_with_format(1, 3, "right bottom", &right_bottom)?;
    // add borders
    let thin_border = Format::default().set_border(FormatBorderType::Double);
    worksheet.write_with_format(2, 1, "bordered text", &thin_border)?;
    let thin_border = Format::default().set_border_bottom(FormatBorderType::Double).set_border_top(FormatBorderType::Double);
    worksheet.write_with_format(2, 3, "bordered text", &thin_border)?;
    let thin_border = Format::default().set_border_right(FormatBorderType::Double).set_border_left(FormatBorderType::Double);
    worksheet.write_with_format(2, 5, "bordered text", &thin_border)?;
    // add background
    let red_background = Format::default().set_background_color(FormatColor::RGB("00FF7777"));
    worksheet.write_with_format(3, 1, "red", &red_background)?;
    // workbook.save_as("examples/xlsx/edit_style.xlsx")?;
    workbook.save()?;
    Ok(())
}