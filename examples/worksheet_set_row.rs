use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    let blue = Format::default().set_color(FormatColor::RGB("007777FF")).set_background_color(FormatColor::RGB("00FF77FF"));
    worksheet.set_row(2, 100.5)?;
    worksheet.set_row_with_format(1, 100.5, &blue);
    workbook.save()?;
    Ok(())
}