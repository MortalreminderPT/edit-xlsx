use edit_xlsx::{Format, Workbook, WorkbookResult, FormatAlignType};
fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::from_path("examples/xlsx/edit.xlsx")?;
    let mut worksheet = workbook.get_worksheet(1)?;
    // adjust the text align
    // let mut binding = Format::default();
    let center = Format::default().set_align(FormatAlignType::VerticalCenter).set_align(FormatAlignType::Center);
    let left_top = Format::default().set_align(FormatAlignType::Left).set_align(FormatAlignType::Top);
    let right_bottom = Format::default().set_align(FormatAlignType::Right).set_align(FormatAlignType::Bottom);
    worksheet.write_with_format(1, 1, "center", &center)?;
    worksheet.write_with_format(1, 2, "left top", &left_top)?;
    worksheet.write_with_format(1, 3, "right bottom", &right_bottom)?;
    workbook.save_as("examples/output/edit_style.xlsx")?;
    Ok(())
}