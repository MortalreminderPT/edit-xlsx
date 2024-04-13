#[cfg(test)]
mod tests {
    use std::thread;
    use edit_xlsx::{Workbook, WorkbookResult};

    #[test]
    fn test_new() -> WorkbookResult<()> {
        let workbook = Workbook::new();
        let mut handles = vec![];
        for i in 1..10 {
            handles.push(thread::spawn(move || {
                let workbook = Workbook::new();
                workbook.save_as(format!("tests/output/test_concurrent_new{i}.xlsx"))?;
                Ok(())
            }));
        }
        let _: Vec<WorkbookResult<()>> = handles
            .into_iter()
            .map(|h| h.join().unwrap())
            .collect();
        workbook.save_as("tests/output/test_concurrent_new.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from() -> WorkbookResult<()> {
        let workbook = Workbook::from_path("tests/xlsx/image_nao.xlsx")?;
        let mut handles = vec![];
        for i in 1..10 {
            handles.push(thread::spawn(move || {
                let workbook = Workbook::from_path("tests/xlsx/image_nao.xlsx")?;
                workbook.save_as(format!("tests/output/test_concurrent_from{i}.xlsx"))?;
                Ok(())
            }));
        }
        let _: Vec<WorkbookResult<()>> = handles
            .into_iter()
            .map(|h| h.join().unwrap())
            .collect();
        workbook.save_as("tests/output/test_concurrent_from.xlsx")?;
        Ok(())
    }
}