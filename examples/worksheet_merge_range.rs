use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor, FormatAlignType};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    let format = Format::default()
        .set_color(FormatColor::RGB("00000000"))
        .set_background_color(FormatColor::RGB("00ff0000"))
        .set_align(FormatAlignType::Center)
        .set_align(FormatAlignType::VerticalCenter);
    let format2 = Format::default()
        .set_color(FormatColor::RGB("00000000"))
        .set_background_color(FormatColor::RGB("00D7E4BC"))
        .set_align(FormatAlignType::Center)
        .set_align(FormatAlignType::VerticalCenter);
    worksheet.set_column(1, 10, 20.5)?;
    worksheet.merge_range(1, 1, 5, 5, "merge cell", &format)?;
    worksheet.merge_range(3, 6, 10, 10, "merge cell", &format2)?;
    workbook.save()?;
    Ok(())
}