#[cfg(test)]
mod tests {
    use edit_xlsx::{Workbook, WorkbookResult};
    
    #[test]
    fn test_new() -> WorkbookResult<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.get_worksheet_mut(1)?;
        workbook.save_as("tests/output/{file_name}_test_new.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("aaa.xlsx")?;
        // let worksheet = workbook.get_worksheet_mut(1)?;
        workbook.save_as("tests/output/rich_text_test_from.xlsx")?;
        Ok(())
    }
}