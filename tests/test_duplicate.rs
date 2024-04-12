#[cfg(test)]
mod tests {
    use edit_xlsx::{Workbook, WorkbookResult, Write};

    #[test]
    fn test_new() -> WorkbookResult<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.get_worksheet(1)?;
        let worksheet = workbook.add_worksheet()?;
        let worksheet = workbook.add_worksheet()?;
        let worksheet = workbook.add_worksheet_by_name("worksheet")?;
        worksheet.write("a1", "hello")?;
        let worksheet = workbook.duplicate_worksheet_by_name("worksheet")?;
        worksheet.write("B2", "hi")?;
        workbook.save_as("tests/output/duplicate_test_new.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/many_sheets.xlsx")?;
        for i in 1..6 {
            workbook.duplicate_worksheet(i)?;
        }
        workbook.save_as("tests/output/duplicate_test_from.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_appraisal_score() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/appraisal_score.xlsx")?;
        let names = ["Details", "Template", "Advanced Project Plan Template"];
        for name in names {
            // let worksheet = workbook.get_worksheet_by_name(name)?;
            workbook.duplicate_worksheet_by_name(name)?;
        }
        // Program will panic here because of the duplicated sheet name
        // workbook.duplicate_worksheet_by_name("Template")?;
        workbook.save_as("tests/output/duplicate_test_from_appraisal_score.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_accounting() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/accounting.xlsx")?;
        let sheet = workbook.duplicate_worksheet(1)?;
        sheet.write_column("C6", &(1..140).into_iter().collect::<Vec<i32>>())?;
        workbook.save_as("tests/output/duplicate_test_from_accounting.xlsx")?;
        Ok(())
    }
}