use edit_xlsx::Workbook;

#[cfg(test)]
mod tests {
    use edit_xlsx::WorkbookResult;
    use crate::Workbook;

    #[test]
    fn test_new() -> WorkbookResult<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.get_worksheet(1)?;
        worksheet.insert_image("B14:E24", &"./examples/pics/capybara.bmp")?;
        worksheet.insert_image("E14:H24", &"examples/pics/rust.png")?;
        workbook.save_as("tests/output/image_test_new.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/image_nao.xlsx")?;
        let worksheet = workbook.get_worksheet(1)?;
        worksheet.insert_image("A16:D25", &"./examples/pics/capybara.bmp")?;
        worksheet.insert_image("A7:D16", &"./examples/pics/rust.png")?;
        workbook.save_as("tests/output/image_test_from.xlsx")?;
        Ok(())
    }
    #[test]
    fn test_from_readonly() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/image_nao.xlsx")?;
        workbook.save_as("tests/output/image_test_from_readonly.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_png() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/image_nao.xlsx")?;
        let worksheet = workbook.get_worksheet(1)?;
        worksheet.insert_image("A7:D16", &"./examples/pics/rust.png")?;
        worksheet.insert_image("A10:D19", &"./examples/pics/rust.png")?;
        workbook.save_as("tests/output/image_test_from_png.xlsx")?;
        Ok(())
    }
}