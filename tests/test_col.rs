#[cfg(test)]
mod tests {
    use edit_xlsx::{Col, Workbook, WorkbookResult};

    #[test]
    fn test_new_default_width() -> WorkbookResult<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.get_worksheet_mut(1)?;
        worksheet.set_default_column(18.5);
        assert_eq!(18.5, worksheet.get_default_column());
        workbook.save_as("tests/output/col_test_new_default_col.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_default_width() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/row_and_col.xlsx")?;
        let worksheet = workbook.get_worksheet_mut(1)?;
        assert_eq!(8.11, worksheet.get_default_column());
        worksheet.set_default_column(30.8);
        assert_eq!(30.8, worksheet.get_default_column());
        workbook.save_as("tests/output/col_test_from_default_col.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_new_custom_width() -> WorkbookResult<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.get_worksheet_mut(1)?;
        // assert_eq!(20.0, worksheet.get_col(1)?);
        worksheet.set_columns_width("A:A", 22.5)?;
        assert_eq!(22.5, worksheet.get_columns_width("A:A")?.get("A:A").unwrap().unwrap());//; [0].2.unwrap());
        workbook.save_as("tests/output/col_test_new_custom_col.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_custom_width() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/row_and_col.xlsx")?;
        let worksheet = workbook.get_worksheet_mut(1)?;
        worksheet.set_columns_width("A:A", 100.0)?;
        assert_eq!(100.0, worksheet.get_columns_width("A:A")?.get("A:A").unwrap().unwrap());
        workbook.save_as("tests/output/col_test_from_custom_col.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_read_custom_width() -> WorkbookResult<()> {
        let workbook = Workbook::from_path("tests/xlsx/accounting.xlsx")?;
        let worksheet = workbook.get_worksheet(1)?;
        let widths = worksheet.get_columns_width("A:A");
        Ok(())
    }
}