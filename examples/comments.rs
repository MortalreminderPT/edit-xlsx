use edit_xlsx::{Workbook, WorkbookResult};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::from_path("examples/comments_test.xlsx")?;
    let worksheet = workbook.get_worksheet(1)?;
    workbook.save_as("examples/comments_test_saved.xlsx")?;
    Ok(())
}