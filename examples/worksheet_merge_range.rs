use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    let format1 = Format::default().set_color(FormatColor::RGB("007777FF")).set_background_color(FormatColor::RGB("00FF0000"));
    let format2 = Format::default().set_color(FormatColor::RGB("00ffffff")).set_background_color(FormatColor::RGB("00006600"));
    let format3 = Format::default().set_color(FormatColor::RGB("00000000")).set_background_color(FormatColor::RGB("00000066"));
    let format4 = Format::default().set_color(FormatColor::RGB("00000000")).set_background_color(FormatColor::RGB("00006666"));
    let format5 = Format::default().set_color(FormatColor::RGB("00ff0000")).set_background_color(FormatColor::RGB("00D7E4BC"));
    worksheet.set_row(2, 100.5)?;
    worksheet.set_row_with_format(1, 100.5, &format1);
    worksheet.set_column_with_format(1, 5, 5.0, &format3);
    worksheet.set_column_with_format(2, 4, 5.0, &format4);
    worksheet.merge_range(1, 1, 5, 4, "aa", &format2);
    worksheet.merge_range(3, 5, 10, 10, "bb", &format5);
    workbook.save()?;
    Ok(())
}