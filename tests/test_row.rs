#[cfg(test)]
mod tests {
    use edit_xlsx::{WorkSheetRow, Workbook, WorkbookResult};

    #[test]
    fn test_new_default_height() -> WorkbookResult<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.get_worksheet_mut(1)?;
        worksheet.set_default_row(18.5);
        assert_eq!(18.5, worksheet.get_default_row());
        workbook.save_as("tests/output/row_test_new_default_row.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_default_height() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/row_and_col.xlsx")?;
        let worksheet = workbook.get_worksheet_mut(1)?;
        assert_eq!(20.0, worksheet.get_default_row());
        worksheet.set_default_row(30.8);
        assert_eq!(30.8, worksheet.get_default_row());
        workbook.save_as("tests/output/row_test_from_default_row.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_new_custom_height() -> WorkbookResult<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.get_worksheet_mut(1)?;
        // assert_eq!(20.0, worksheet.get_row(1)?);
        worksheet.set_row_height(1, 22.5)?;
        assert_eq!(22.5, worksheet.get_row_height(1)?);
        workbook.save_as("tests/output/row_test_new_custom_row.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_custom_height() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/row_and_col.xlsx")?;
        let worksheet = workbook.get_worksheet_mut(1)?;
        worksheet.set_row_height(1, 100.0)?;
        assert_eq!(100.0, worksheet.get_row_height(1)?);
        workbook.save_as("tests/output/row_test_from_custom_row.xlsx")?;
        Ok(())
    }
}