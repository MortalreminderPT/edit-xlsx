#[cfg(test)]
mod tests {
    use std::thread;
    use edit_xlsx::{Workbook, WorkbookResult};

    #[test]
    fn test_new() -> WorkbookResult<()> {
        let workbook = Workbook::new();
        let handle = thread::spawn(|| {
            for i in 1..10 {
                let workbook = Workbook::new();
                workbook.save_as(format!("tests/output/test_concurrent_new{i}.xlsx")).unwrap();
            }
        });
        workbook.save_as("tests/output/test_concurrent_new.xlsx")?;
        handle.join().unwrap();
        Ok(())
    }

    #[test]
    fn test_from() -> WorkbookResult<()> {
        let workbook = Workbook::from_path("tests/xlsx/image_nao.xlsx")?;
        let handle = thread::spawn(|| {
            for i in 1..10 {
                let workbook = Workbook::from_path("tests/xlsx/image_nao.xlsx").unwrap();
                workbook.save_as(format!("tests/output/test_concurrent_from{i}.xlsx")).unwrap();
            }
        });
        workbook.save_as("tests/output/test_concurrent_from.xlsx")?;
        handle.join().unwrap();
        Ok(())
    }
}