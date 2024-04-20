#[cfg(test)]
mod tests {
    use edit_xlsx::{FormatColor, FormatFont, Read, RichText, Word, Workbook, WorkbookResult, Write};

    #[test]
    fn test_new() -> WorkbookResult<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.get_worksheet_mut(1)?;
        workbook.save_as("tests/output/{file_name}_test_new.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/accounting.xlsx")?;
        let worksheet = workbook.get_worksheet_mut(1)?;

        let mut font1 = FormatFont::default();
        font1.bold = true;
        let mut font2 = FormatFont::default();
        font2.italic = true;
        font2.color = FormatColor::Index(14);
        let rich_text = RichText::new_word("Hello", &font1) + Word::new("World", &font2);
        println!("{:?}", rich_text);
        worksheet.write_rich_string("A1", &rich_text)?;
        worksheet.write_rich_string("A2", &rich_text)?;
        workbook.save_as("tests/output/rich_text_test_from.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_read_from() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/image_nao.xlsx")?;
        let worksheet = workbook.get_worksheet_mut(1)?;
        let cell = worksheet.read_cell("A1")?;
        let rt = cell.rich_text.unwrap();
        println!("{:?}", rt);
        Ok(())
    }
}