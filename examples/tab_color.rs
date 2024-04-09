use edit_xlsx::{FormatColor, Workbook, WorkbookResult};

fn main() -> WorkbookResult<()> {
    // Create a new workbook
    let mut workbook = Workbook::new();
    // Set up some worksheets and set tab colors
    let worksheet1 = workbook.get_worksheet(1)?;
    worksheet1.set_tab_color(&FormatColor::RGB("00ff0000".to_string())); // Red
    let worksheet2 = workbook.add_worksheet()?;
    worksheet2.set_tab_color(&FormatColor::Index(11)); // Green
    let worksheet3 = workbook.add_worksheet()?;
    worksheet3.set_tab_color(&FormatColor::Theme(5, 0.0)); // Orange
    let worksheet4 = workbook.add_worksheet()?;
    // worksheet4 will have the default color.
    workbook.save_as("examples/tab_color.xlsx")?;
    Ok(())
}