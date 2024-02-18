use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    // string
    worksheet.write_string(1, 1, "123456".to_string())?;
    // number
    worksheet.write_number(2, 1, 123456)?;
    worksheet.write_double(3, 1, 3.14159265)?;
    // formula
    worksheet.write_number(1, 3, 1)?;
    worksheet.write_number(1, 4, 2)?;
    worksheet.write_number(2, 3, 3)?;
    worksheet.write_number(2, 4, 4)?;
    worksheet.write_formula(4, 1, "=COS(2/3*PI())")?;
    worksheet.write_formula(5, 1, "=C1 + D2")?;
    worksheet.write_formula(6, 1, "=SUM(B1:B5)")?;
    worksheet.write_formula(7, 1, "=IF(A3>1,\"Yes\", \"No\")")?;
    worksheet.write_formula(8, 1, "=AVERAGE(1, 2, 3, 4)")?;
    worksheet.write_formula(9, 1, "=DATEVALUE(\"1-Jan-2013\")")?;
    worksheet.write_array_formula(10, 1, "=SUM(C1:D1*C2:D2)")?;
    worksheet.write_dynamic_array_formula("B11:B13", "LEN(A2:A4)")?;
    // url
    worksheet.write_url(11, 1, "https://www.rust-lang.org/")?;
    worksheet.write_url_with_text(12, 1, "https://www.rust-lang.org/", "rust")?;
    workbook.save()?;
    Ok(())
}