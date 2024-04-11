#[cfg(test)]
mod tests {
    use edit_xlsx::{Col, Row, Workbook, WorkbookResult};

    #[test]
    fn test_new() -> WorkbookResult<()> {
        let mut workbook = Workbook::new();
        let collapse_row_sheet = workbook.get_worksheet(1)?;
        collapse_row_sheet.set_name("Row Collapse")?;
        for row in 1..20 { collapse_row_sheet.set_row_level(row, 1)?; }
        for row in 5..15 { collapse_row_sheet.set_row_level(row, 2)?; }
        for row in 8..13 { collapse_row_sheet.set_row_level(row, 3)?; }
        for row in 9..10 {
            collapse_row_sheet.set_row_level(row, 4)?;
            collapse_row_sheet.hide_row(row)?;
        }
        let collapse_col_sheet = workbook.add_worksheet()?;
        collapse_col_sheet.set_name("Col Collapse")?;
        collapse_col_sheet.set_column_level("A:Z", 1)?;
        // todo: Add a method to read column default col width
        collapse_col_sheet.set_column("A:Z", 20.0)?;
        collapse_col_sheet.set_column_level("E:Q", 2)?;
        collapse_col_sheet.set_column_level("G:J", 3)?;
        collapse_col_sheet.set_column_level("H:I", 4)?;
        collapse_col_sheet.set_column("H:I", 0.0)?;
        workbook.save_as("tests/output/outline_collapse_test_new.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_new_both() -> WorkbookResult<()> {
        let mut workbook = Workbook::new();
        let collapse_row_sheet = workbook.get_worksheet(1)?;
        collapse_row_sheet.set_name("Row Collapse")?;
        for row in 1..20 { collapse_row_sheet.set_row_level(row, 1)?; }
        for row in 5..15 { collapse_row_sheet.set_row_level(row, 2)?; }
        for row in 8..13 { collapse_row_sheet.set_row_level(row, 3)?; }
        for row in 9..10 {
            collapse_row_sheet.set_row_level(row, 4)?;
            collapse_row_sheet.hide_row(row)?;
        }
        collapse_row_sheet.deselect();
        let both_sheet = workbook.duplicate_worksheet_by_name("Row Collapse")?;
        both_sheet.set_name("Col Collapse")?;
        both_sheet.set_column_level("A:Z", 1)?;
        both_sheet.set_column("A:Z", 20.0)?;
        both_sheet.set_column_level("E:Q", 2)?;
        both_sheet.set_column_level("G:J", 3)?;
        both_sheet.set_column_level("H:I", 4)?;
        both_sheet.set_column("H:I", 0.0)?;
        both_sheet.activate();
        workbook.save_as("tests/output/outline_collapse_test_new_both.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/background_capybara.xlsx")?;
        let collapse_row_sheet = workbook.get_worksheet(1)?;
        collapse_row_sheet.set_name("Row Collapse")?;
        for row in 1..20 { collapse_row_sheet.set_row_level(row, 1)?; }
        for row in 5..15 { collapse_row_sheet.set_row_level(row, 2)?; }
        for row in 8..13 { collapse_row_sheet.set_row_level(row, 3)?; }
        for row in 9..10 {
            collapse_row_sheet.set_row_level(row, 4)?;
            collapse_row_sheet.hide_row(row)?;
        }
        let collapse_col_sheet = workbook.add_worksheet()?;
        collapse_col_sheet.set_name("Col Collapse")?;
        collapse_col_sheet.set_column_level("A:Z", 1)?;
        // todo: Add a method to read column default col width
        collapse_col_sheet.set_column("A:Z", 20.0)?;
        collapse_col_sheet.set_column_level("E:Q", 2)?;
        collapse_col_sheet.set_column_level("G:J", 3)?;
        collapse_col_sheet.set_column_level("H:I", 4)?;
        collapse_col_sheet.set_column("H:I", 0.0)?;
        workbook.save_as("tests/output/outline_collapse_test_from.xlsx")?;
        Ok(())
    }
    
    #[test]
    fn test_from_with_image() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/image_nao.xlsx")?;
        let collapse_row_sheet = workbook.get_worksheet(1)?;
        collapse_row_sheet.set_name("Row Collapse")?;
        for row in 1..20 { collapse_row_sheet.set_row_level(row, 1)?; }
        for row in 5..15 { collapse_row_sheet.set_row_level(row, 2)?; }
        for row in 8..13 { collapse_row_sheet.set_row_level(row, 3)?; }
        for row in 9..10 {
            collapse_row_sheet.set_row_level(row, 4)?;
            collapse_row_sheet.hide_row(row)?;
        }
        collapse_row_sheet.deselect();
        let both_sheet = workbook.duplicate_worksheet_by_name("Row Collapse")?;
        both_sheet.set_name("Col Collapse")?;
        both_sheet.set_column_level("A:Z", 1)?;
        both_sheet.set_column("A:Z", 20.0)?;
        both_sheet.set_column_level("E:Q", 2)?;
        both_sheet.set_column_level("G:J", 3)?;
        both_sheet.set_column_level("H:I", 4)?;
        both_sheet.set_column("H:I", 0.0)?;
        both_sheet.activate();
        let image_sheet = workbook.duplicate_worksheet_by_name("Col Collapse")?;
        image_sheet.insert_image("B14:E24", &"./examples/pics/capybara.bmp")?;
        image_sheet.insert_image("E14:H24", &"examples/pics/rust.png")?;
        workbook.save_as("tests/output/outline_collapse_test_from_with_image.xlsx")?;
        Ok(())
    }
}