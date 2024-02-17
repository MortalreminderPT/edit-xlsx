use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    let blue = Format::default().set_color(FormatColor::RGB("007777FF")).set_background_color(FormatColor::RGB("00FF77FF"));
    worksheet.set_column(1, 5, 10.0);
    worksheet.write(1, 1, "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaa");
    worksheet.write(1, 3, "bbbbbbbbbb");
    worksheet.write(1, 5, "cccccccccccccccc");
    worksheet.write(1, 2, "ddddddddddddddddddddddddddddddddddd");
    let worksheet = workbook.add_worksheet_by_name("你好")?;
    let worksheet = workbook.add_worksheet_by_name("你好2")?;
    let worksheet = workbook.add_worksheet_by_name("你好3")?;
    let worksheet = workbook.add_worksheet_by_name("你好4")?;
    for worksheet in workbook.worksheets() {
        worksheet.select();
        let a = worksheet.get_name();
        println!("{}", a);
    }
    workbook.get_worksheet_by_name("你好3")?.set_first_sheet();
    workbook.save()?;
    Ok(())
}