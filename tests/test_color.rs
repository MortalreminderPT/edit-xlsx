#[cfg(test)]
mod tests {
    use edit_xlsx::{WorkSheetCol, Format, FormatColor, Row, Workbook, WorkbookResult, Write};

    #[test]
    fn test_new() -> WorkbookResult<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.get_worksheet_mut(1)?;
        // rgb
        let red = Format::default().set_background_color(FormatColor::RGB(255, 0, 0));
        // index
        let green = Format::default().set_background_color(FormatColor::Index(11));
        // theme
        let blue = Format::default().set_background_color(FormatColor::Theme(4, 0.0));
        // cell color
        worksheet.merge_range("A1:C1", "TEST CELL COLOR")?;
        worksheet.write_with_format("A2", "RED", &red)?;
        worksheet.write_with_format("B2", "GREEN", &green)?;
        worksheet.write_with_format("C2", "BLUE", &blue)?;
        // col color
        worksheet.merge_range("D1:F1", "TEST COLUMN COLOR")?;
        worksheet.set_columns_width_with_format("D:D", 10.0, &red)?;
        worksheet.set_columns_width_with_format("E:E", 10.0, &green)?;
        worksheet.set_columns_width_with_format("F:F", 10.0, &blue)?;
        worksheet.write("D2", "RED")?;
        worksheet.write("E2", "GREEN")?;
        worksheet.write("F2", "BLUE")?;
        // row color
        worksheet.merge_range((5, 1, 5, 16384), "TEST ROW COLOR")?;
        worksheet.set_row_with_format(6, 15.0, &red)?;
        worksheet.set_row_with_format(7, 15.0, &green)?;
        worksheet.set_row_with_format(8, 15.0, &blue)?;
        worksheet.write("A6", "RED")?;
        worksheet.write("A7", "GREEN")?;
        worksheet.write("A8", "BLUE")?;
        workbook.save_as("tests/output/color_test_new.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/color.xlsx")?;
        let worksheet = workbook.get_worksheet_mut(1)?;
        // cell color
        worksheet.write("A1", "RED")?;
        worksheet.write("B1", "GREEN")?;
        worksheet.write("C1", "BLUE")?;
        let mut text_red = Vec::new();
        let mut text_green = Vec::new();
        let mut text_blue = Vec::new();
        for i in 1..=10 {
            text_red.push(format!("RED{i}"));
            text_green.push(format!("GREEN{i}"));
            text_blue.push(format!("BLUE{i}"));
        }
        // col color
        worksheet.write_column("D1", &text_red)?;
        worksheet.write_column("E1", &text_green)?;
        worksheet.write_column("F1", &text_blue)?;
        // row color
        worksheet.write_row("A6", &text_red)?;
        worksheet.write_row("A7", &text_green)?;
        worksheet.write_row("A8", &text_blue)?;
        workbook.save_as("tests/output/color_test_from.xlsx")?;
        Ok(())
    }
}