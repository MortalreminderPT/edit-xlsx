use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    let tab_color = FormatColor::RGB("00000000");
    worksheet.set_tab_color(&tab_color);
    workbook.save()?;
    Ok(())
}