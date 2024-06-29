#[cfg(test)]
mod tests {
    use edit_xlsx::{Read, Workbook, WorkbookResult};

    #[test]
    fn test_new() -> WorkbookResult<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.get_worksheet_mut(1)?;
        workbook.save_as("tests/output/{file_name}_test_new.xlsx")?;
        Ok(())
    }

    #[test]
    #[cfg(feature = "ansi_term_support")]
    fn test_from() -> WorkbookResult<()> {
        use ansi_term::ANSIStrings;
        let mut workbook = Workbook::from_path("tests/xlsx/accounting.xlsx")?;
        let worksheet = workbook.get_worksheet_mut(1)?;
        let theme = worksheet.get_theme(0);
        let cell = worksheet.read_cell("A1")?;
        // let display = theme.ansi_strings(&cell);
        // print!("{}", ANSIStrings(&display));
        workbook.save_as("tests/output/{file_name}_test_from.xlsx")?;
        Ok(())
    }

    #[test]
    #[cfg(feature = "ansi_term_support")]
    fn test_read_ansi_from() -> WorkbookResult<()> {
        use ansi_term::ANSIStrings;
        let reading_book = Workbook::from_path("./tests/xlsx/appraisal_score.xlsx")?;
        let reading_sheet = reading_book.get_worksheet_by_name("Details")?;
        // Read then write text and format
        for row in 1..=reading_sheet.max_column() {
            for col in 1..=reading_sheet.max_row() {
                print!("{}\t", ANSIStrings(&reading_sheet.ansi_strings((row, col))?));
            }
            println!();
        }
        print!("{}\t", ANSIStrings(&reading_sheet.ansi_strings("C12")?));
        Ok(())
    }

    #[test]
    fn test_beautify() -> WorkbookResult<()> {
        #[cfg(feature = "ansi_term_support")]
        {
            // use ansi_term::{ANSIStrings};
            // let workbook = Workbook::from_path("tests/xlsx/paycheck-calculator.xlsx")?;
            // let worksheet = workbook.get_worksheet_by_name("NEW W-4")?;
            // let cell = worksheet.read_cell("A2")?;
            // println!("{}", ANSIStrings(&cell.ansi_strings()));
            // let workbook = Workbook::from_path("tests/xlsx/rich-text.xlsx")?;
            // let worksheet = workbook.get_worksheet_by_name("Sheet1")?;
            // let cell = worksheet.read_cell("A1")?;
            // println!("{}", ANSIStrings(&cell.ansi_strings()));
        }
        Ok(())
    }
}