#[cfg(test)]
mod tests {
    use edit_xlsx::{Workbook, WorkbookResult, Write};

    #[test]
    fn test_new_write() -> WorkbookResult<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.get_worksheet_mut(1)?;
        // test write cell
        worksheet.write("A1", "Text")?;
        let texts: Vec<String> = (1..=100).into_iter().map(|i| format!("Text{i}")).collect();
        // test write row
        worksheet.write_row("A2", &texts)?;
        // test write col
        worksheet.write_column("A3", &texts)?;
        workbook.save_as("tests/output/text_test_new_write.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_write_analysis_tabs() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/appraisal_score.xlsx")?;
        let worksheet = workbook.get_worksheet_by_name("Template")?;
        let texts: Vec<String> = (1..=100).into_iter().map(|i| format!("Text{i}")).collect();
        // test write col
        worksheet.write_column("B2", &texts)?;
        workbook.save_as("tests/output/test_from_write_analysis_tabs.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_write() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/appraisal_score.xlsx")?;
        let worksheet = workbook.get_worksheet_by_name("Template")?;
        // test write cell
        worksheet.write("P1", "Text")?;
        let texts: Vec<String> = (1..=100).into_iter().map(|i| format!("Text{i}")).collect();
        let nums: Vec<i32> = (-100..=-1).into_iter().collect();
        // test write row
        worksheet.write_row("P1", &texts)?;
        // test write col
        worksheet.write_column("C2", &texts)?;
        worksheet.write_column("D2", &nums)?;
        workbook.save_as("tests/output/text_test_from_write.xlsx")?;
        Ok(())
    }
}