use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor, Write, FormatAlignType};

fn main() -> WorkbookResult<()> {
    let format = Format::default()
        .set_color(FormatColor::RGB("00000000"))
        .set_background_color(FormatColor::RGB("00ff0000"))
        .set_align(FormatAlignType::Center)
        .set_align(FormatAlignType::VerticalCenter);
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    // write string
    worksheet.write_string("A1", "string".to_string())?;
    // write number
    worksheet.write_number("A2", 1024)?;
    worksheet.write_double("A3", std::f64::consts::PI)?;
    // write formula
    worksheet.write_formula("4A", "=COS(2/3*PI())")?;
    worksheet.write_formula("4B", "=B1 + C2")?;
    worksheet.write_formula("4C", "=SUM(B1:B5)")?;
    worksheet.write_formula((4, 4), "=IF(A3>1,\"Yes\", \"No\")")?;
    worksheet.write_formula((4, 5), "=AVERAGE(1, 2, 3, 4)")?;
    worksheet.write_formula((4, 6), "=DATEVALUE(\"1-Jan-2013\")")?;
    worksheet.write_number((1, 2), 1)?;
    worksheet.write_number((1, 3), 20)?;
    worksheet.write_number((2, 2), 3)?;
    worksheet.write_number((2, 3), 4)?;
    worksheet.write_array_formula((4, 7), "=SUM(B1:C1*B2:C2)")?;
    worksheet.write_dynamic_array_formula("H4:J4", "LEN(A1:C1)")?;
    // write url
    worksheet.write_url((5, 1), "https://www.rust-lang.org/")?;
    worksheet.write_url_text((5, 2), "https://www.rust-lang.org/", "rust")?;
    // write row
    let data = vec![1, 2, 3, 4, 5];
    worksheet.write_row_with_format((6, 5), data, &format)?;
    // write column
    let data = vec!["col_1", "col_2", "col_3", "col_4", "col_5", "col_6", "col_7", "col_8", "col_9", "col_10"];
    worksheet.write_column((7, 1), data)?;
    workbook.save()?;

    Ok(())
}