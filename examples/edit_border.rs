use edit_xlsx::{Format, Workbook, WorkbookResult, FormatBorderType};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::from_path("examples/xlsx/edit.xlsx")?;
    let mut worksheet = workbook.get_worksheet(1)?;
    // add borders
    let thin_border = Format::default().set_border(FormatBorderType::Double);
    worksheet.write_with_format(1, 1, "bordered text", &thin_border)?;
    let thin_border = Format::default().set_border_bottom(FormatBorderType::Double).set_border_top(FormatBorderType::Double);
    worksheet.write_with_format(1, 3, "bordered text", &thin_border)?;
    let thin_border = Format::default().set_border_right(FormatBorderType::Double).set_border_left(FormatBorderType::Double);
    worksheet.write_with_format(1, 5, "bordered text", &thin_border)?;
    workbook.save_as("examples/output/edit_border.xlsx")?;
    Ok(())
}