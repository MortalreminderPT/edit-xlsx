#[cfg(test)]
mod tests {
    use edit_xlsx::{Workbook, WorkbookResult};

    #[test]
    fn test_new() -> WorkbookResult<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.get_worksheet_mut(1)?;
        worksheet.set_background("examples/pics/ferris.png")?;
        workbook.save_as("tests/output/background_test_new.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_new_overwrite() -> WorkbookResult<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.get_worksheet_mut(1)?;
        worksheet.set_background("examples/pics/ferris.png")?;
        worksheet.set_background("examples/pics/rust.png")?;
        workbook.save_as("tests/output/background_test_new_overwrite.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_readonly() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/background_capybara.xlsx")?;
        let worksheet = workbook.get_worksheet_mut(1)?;
        workbook.save_as("tests/output/background_test_from_readonly.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_overwrite() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/background_capybara.xlsx")?;
        let worksheet = workbook.get_worksheet_mut(1)?;
        worksheet.set_background("examples/pics/ferris.png")?;
        workbook.save_as("tests/output/background_test_from_overwrite.xlsx")?;
        Ok(())
    }
}