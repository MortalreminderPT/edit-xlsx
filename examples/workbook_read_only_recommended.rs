use edit_xlsx::{Workbook, WorkbookResult};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    workbook.set_size(1200, 800)?;
    workbook.read_only_recommended()?;
    workbook.save()?;
    Ok(())
}