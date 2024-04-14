#[cfg(test)]
mod tests {
    use edit_xlsx::{Cell, Format, FormatColor, Workbook, WorkbookResult, Write};

    #[test]
    fn test_new() -> WorkbookResult<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.get_worksheet_mut(1)?;
        workbook.save_as("tests/output/cell_test_new.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_write_analysis_tabs() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/appraisal_score.xlsx")?;
        let worksheet = workbook.get_worksheet_mut_by_name("Template")?;
        let texts: Vec<String> = (1..=100).into_iter().map(|i| format!("Text{i}")).collect();
        let mut cell = Cell::default();
        cell.text = Some(123456);
        cell.format = Some(Format::default().set_font("Elephant").set_underline().set_color(FormatColor::RGB(0, 0, 255)));
        cell.hyperlink = Some("https://google.com.tw".to_string());
        // cell.formula = Some("SUM(D2:E26)".to_string());
        // test write col
        worksheet.write_cell("B4", &cell)?;
        workbook.save_as("tests/output/test_from_write_cell_analysis_tabs.xlsx")?;
        Ok(())
    }
}