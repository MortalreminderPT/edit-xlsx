use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor, Write, Col, FormatAlignType};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    let blue = Format::default()
        .set_color(FormatColor::RGB("007777FF"))
        .set_background_color(FormatColor::RGB("000077FF"))
        .set_align(FormatAlignType::VerticalCenter)
        .set_align(FormatAlignType::Center);
    worksheet.set_column_with_format(1, 5, 100.0, &blue)?;
    worksheet.write_with_format(1, 1, "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", &blue)?;
    worksheet.write_with_format(1, 3, "bbbbbbbbbb", &blue)?;
    // worksheet.write(1, 5, "cccccccccccccccc")?;
    worksheet.write_with_format(1, 2, "ddddddddddddddddddddddddddddddddddd", &blue)?;
    let worksheet = workbook.add_worksheet_by_name("你好")?;
    let worksheet = workbook.add_worksheet_by_name("你好2")?;
    let worksheet = workbook.add_worksheet_by_name("你好3")?;
    let worksheet = workbook.add_worksheet_by_name("你好4")?;
    for worksheet in workbook.worksheets() {
        worksheet.select();
        let a = worksheet.get_name();
        println!("{}", a);
    }
    workbook.get_worksheet_by_name("你好3")?; //.set_first_sheet();
    workbook.save()?;
    Ok(())
}