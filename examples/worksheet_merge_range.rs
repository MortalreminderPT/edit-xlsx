use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor, FormatAlignType, Col, Write};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    let format = Format::default()
        .set_color(FormatColor::RGB("00ffffff"))
        .set_background_color(FormatColor::RGB("00ff0000"))
        .set_align(FormatAlignType::Center)
        .set_align(FormatAlignType::VerticalCenter);
    let format2 = Format::default()
        .set_color(FormatColor::RGB("00ffffff"))
        .set_background_color(FormatColor::RGB("00D7E4BC"))
        .set_align(FormatAlignType::Center)
        .set_align(FormatAlignType::VerticalCenter);
    worksheet.set_column(1, 10, 20.5)?;
    worksheet.merge_range_with_format("2A:B3", "merge cell", &format)?;
    worksheet.merge_range_with_format("5A:10C", "merge cell", &format2)?;
    workbook.save()?;
    Ok(())
}