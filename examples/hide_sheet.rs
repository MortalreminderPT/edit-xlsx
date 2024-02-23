use edit_xlsx::{Col, Workbook, WorkbookResult, Write};

fn main() -> WorkbookResult<()> {
    // Create a new workbook
    let mut workbook = Workbook::new();
    let worksheet1 = workbook.get_worksheet(1)?;

    worksheet1.set_column("A:A", 30.0)?;
    worksheet1.write("A1", "Sheet2 is hidden")?;

    // Hide Sheet2. It won't be visible until it is unhidden in Excel.
    let worksheet2 = workbook.add_worksheet()?;
    worksheet2.set_column("A:A", 30.0)?;
    // worksheet2.activate();
    worksheet2.hide();
    worksheet2.write("A1", "Now it's my turn to find you!")?;
    // Note, you can't hide the "active" worksheet, which generally is the
    // first worksheet, since this would cause an Excel error. So, in order to hide
    // the first sheet you will need to activate another worksheet:
    //
    //    worksheet2.activate();
    //    worksheet1.hide();

    let worksheet3 = workbook.add_worksheet()?;
    worksheet3.set_column("A:A", 30.0)?;
    worksheet3.write("A1", "Sheet2 is hidden")?;

    workbook.save_as("examples/hide_sheet.xlsx")?;
    Ok(())
}