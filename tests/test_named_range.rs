#[cfg(test)]
mod tests {
    use edit_xlsx::{Workbook, WorkbookResult};

    /* Test New is not here, as the examples and the test_from handles it */

    #[test]
    fn test_from() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/named_range_test_from.xlsx")?;
        let global_single = ("GlobalSingle", "Sheet1!$A$1");
        let global_multi = ("GlobalMultispan", "Sheet1!$A$2:$D$2");
        let local_single1 = ("LocalSingle", "Sheet1!$B$6");
        let local_multi1 = ("LocalMultispan", "Sheet1!$B$7:$G$8");
        let local_single2 = ("LocalSingle", "Sheet2!$A$1");
        let local_multi2 = ("LocalMultispan", "Sheet2!$A$2:$F$3");
        let local_new = ("LocalNew", "Sheet2!$B$2:$D$4");
        let global_new = ("GlobalNew", "Sheet1!$D$5");
        // Add some new named ranges
        workbook.define_name(global_new.0, global_new.1).unwrap();
        workbook
            .define_local_name(local_new.0, local_new.1, 2)
            .unwrap();

        let run = |sheet: Option<u32>, pair: (&str, &str)| {
            let found = if let Some(sheet) = sheet {
                workbook.get_defined_local_name(pair.0, sheet)
            } else {
                workbook.get_defined_name(pair.0)
            };
            let found = found.as_ref().map(|s| s.as_str());
            assert_eq!(found.ok(), Some(pair.1));
        };

        run(None, global_single);
        run(None, global_multi);
        run(Some(1), local_single1);
        run(Some(1), local_multi1);
        run(Some(2), local_single2);
        run(Some(2), local_multi2);
        // Check the new items
        run(Some(2), local_new);
        run(None, global_new);
        workbook.save_as("tests/output/named_range_test_from.xlsx")?;
        Ok(())
    }
}
