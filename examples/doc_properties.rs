use edit_xlsx::{Col, Properties, Workbook, WorkbookResult, Write};

fn main() -> WorkbookResult<()> {
    // Create a new workbook
    let mut workbook = Workbook::new();
    let mut properties = Properties::default();
    properties.set_title("This is an example spreadsheet")
        .set_subject("With document properties")
        .set_author("pt")
        .set_manager("example manager")
        .set_company("example company")
        .set_category("Example spreadsheets")
        .set_keywords("Sample, Example, Properties")
        .set_comments("Created with Rust")
        .set_status("example status");
    workbook.set_properties(&properties)?;
    // Use the default worksheet
    let worksheet = workbook.get_worksheet(1)?;
    worksheet.set_column("A:A", 70.0)?;
    worksheet.write("A1", "Select 'Workbook Properties' to see properties.")?;
    workbook.save_as("examples/doc_properties.xlsx")?;
    Ok(())
}