#[cfg(test)]
mod tests {
    use edit_xlsx::{Format, FormatColor, FormatFont, Read, RichText, Word, Workbook, WorkbookResult, Write};

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
    fn test_from_cell() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/rich-text.xlsx")?;
        let worksheet = workbook.get_worksheet_mut(1)?;
        let cell = worksheet.read_cell("A2")?;
        println!("{:#?}", cell.format.unwrap().font.color);
        println!("{:#?}", cell.rich_text.unwrap());
        let cell = worksheet.read_cell("A3")?;
        println!("{:#?}", cell.format.unwrap().font.color);
        println!("{:#?}", cell.rich_text.unwrap());
        Ok(())
    }

    #[test]
    fn test_read_words() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/rich-text.xlsx")?;
        let worksheet = workbook.get_worksheet_mut(1)?;
        let mut cell = worksheet.read_cell("A1")?;
        let mut rich_text = cell.rich_text.unwrap_or_default();
        let rich_text2 = RichText::new_word("Hello", &FormatFont::default());
        let mut red = FormatFont::default();
        red.color = FormatColor::RGB(255, 0, 0);
        let rich_text3 = RichText::new_word(" World", &red);
        rich_text = rich_text + &rich_text2 + &rich_text3;
        println!("{:?}", rich_text);
        println!("{:?}", rich_text2);
        println!("{:?}", rich_text3);
        rich_text.words.iter().for_each(|w| print!("{}", w.text));
        let mut format_font = FormatFont::default();
        format_font.bold = true;
        format_font.italic = false;
        format_font.underline = true;
        rich_text.words.iter_mut().for_each(|w| {
            w.font = Some(format_font.clone())
        });
        cell.rich_text = Some(rich_text);
        worksheet.write_cell("A1", &cell)?;
        workbook.save_as("tests/output/rich_text_test_read_words.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_read_from() -> WorkbookResult<()> {
        let workbook = Workbook::from_path("tests/xlsx/rich-text.xlsx")?;
        let worksheet = workbook.get_worksheet(1)?;
        let cell = worksheet.read_cell("A1")?;
        let rt = cell.rich_text.unwrap();
        println!("{:?}", rt);
        let mut workbook = Workbook::from_path("tests/xlsx/image_nao.xlsx")?;
        let worksheet = workbook.get_worksheet_mut(1)?;
        let mut cell = worksheet.read_cell("A1")?;
        let mut rt = cell.rich_text.unwrap();
        let mut font = FormatFont::default();
        font.italic = true;
        font.size = 30.0;
        font.color = FormatColor::Index(12);
        rt = rt + Word::new(" World", &font);
        println!("{:?}", rt);
        cell.rich_text = Some(rt);
        worksheet.write_cell("A1", &cell)?;
        #[cfg(feature = "ansi_term_support")]
        {
            use ansi_term::{ANSIStrings};
            println!("{}", ANSIStrings(&cell.ansi_strings()));
        }
        workbook.save_as("tests/output/rich_text_test_read_from.xlsx")?;
        Ok(())
    }


    #[test]
    fn test_beautify() -> WorkbookResult<()> {
        #[cfg(feature = "ansi_term_support")]
        {
            use ansi_term::{ANSIStrings};
            let workbook = Workbook::from_path("tests/xlsx/paycheck-calculator.xlsx")?;
            let worksheet = workbook.get_worksheet_by_name("NEW W-4")?;
            let cell = worksheet.read_cell("A2")?;
            cell.ansi_strings();
            println!("{}", ANSIStrings(&cell.ansi_strings()));
            let workbook = Workbook::from_path("tests/xlsx/rich-text.xlsx")?;
            let worksheet = workbook.get_worksheet_by_name("Sheet1")?;
            let cell = worksheet.read_cell("A1")?;
            cell.ansi_strings();
            println!("{}", ANSIStrings(&cell.ansi_strings()));
        }
        Ok(())
    }
}