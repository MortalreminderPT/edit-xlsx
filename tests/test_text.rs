#[cfg(test)]
mod tests {
    use edit_xlsx::{Workbook, WorkbookResult, Write};

    #[test]
    fn test_new_write() -> WorkbookResult<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.get_worksheet(1)?;
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
        let mut workbook = Workbook::from_path("tests/xlsx/ANALYSISTABS.xlsx")?;
        // let worksheet = workbook.get_worksheet(1)?;
        // let texts: Vec<String> = (1..=100).into_iter().map(|i| format!("Text{i}")).collect();
        // // test write col
        // worksheet.write_column("B2", &texts)?;
        workbook.save_as("tests/output/test_from_write_analysis_tabs.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_write() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/ANALYSISTABS.xlsx")?;
        let worksheet = workbook.get_worksheet(1)?;
        // test write cell
        worksheet.write("A1", "Text")?;
        let texts: Vec<String> = (1..=100).into_iter().map(|i| format!("Text{i}")).collect();
        // test write row
        worksheet.write_row("A2", &texts)?;
        // test write col
        worksheet.write_column("A3", &texts)?;
        workbook.save_as("tests/output/text_test_from_write.xlsx")?;
        Ok(())
    }
}