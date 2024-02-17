use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    let format = Format::default()
        .set_color(FormatColor::RGB("00000000"))
        .set_background_color(FormatColor::RGB("00D7E4BC"));
    worksheet.set_row(2, 100.5)?;
    worksheet.merge_range(3, 5, 10, 10, "merge cell", &format);
    workbook.save()?;
    Ok(())
}