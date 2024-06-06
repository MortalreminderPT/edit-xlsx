#[cfg(test)]
mod tests {
    use std::num::ParseIntError;
    use edit_xlsx::{Workbook, WorkbookResult, WorkbookError};

    #[test]
    fn test_new() {
        fn test_parse() -> Result<(), MyError> {
            let i = "10";
            let mut workbook = Workbook::new();
            let worksheet = workbook.get_worksheet_mut(i.parse()?)?;
            Ok(())
        }
        test_parse().unwrap_err();
    }

    #[test]
    fn test_from() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/{file_name}.xlsx")?;
        let worksheet = workbook.get_worksheet_mut(1)?;
        workbook.save_as("tests/output/{file_name}_test_from.xlsx")?;
        Ok(())
    }

    #[derive(Debug)]
    enum MyError {
        ParseIntError(ParseIntError),
        WorkbookError(WorkbookError),
    }

    impl From<WorkbookError> for MyError {
        fn from(value: WorkbookError) -> Self {
            MyError::WorkbookError(value)
        }
    }
    impl From<ParseIntError> for MyError {
        fn from(value: ParseIntError) -> Self {
            MyError::ParseIntError(value)
        }
    }
}
