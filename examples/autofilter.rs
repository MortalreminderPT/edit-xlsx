use std::fs;
use edit_xlsx::{Col, Filter, Filters, Format, Row, Workbook, WorkbookResult, Write};

fn main() -> WorkbookResult<()> {
    let text = fs::read_to_string("examples/autofilter_data.txt").unwrap();
    let mut text = text.split("\n");
    let headers: Vec<&str> = text.next().unwrap().split_whitespace().collect();
    let mut data = vec![];
    for text in text {
        let row: Vec<&str> = text.split_whitespace().collect();
        data.push(row);
    }

    // Create a new workbook
    let mut workbook = Workbook::new();
    // Use the default worksheet

    // Set up several sheets with the same data.
    for worksheet in workbook.worksheets() {
        // Make the columns wider.
        worksheet.set_column("A:D", 12.0)?;
        // // Make the header row larger.
        worksheet.set_row_with_format(1, 20.0, &Format::default().set_bold())?;
        // Make the headers bold.
        worksheet.write_row("A1", headers.iter())?;
    }

    let worksheet1 = workbook.get_worksheet(1)?;
    worksheet1.autofilter("A1:D51");

    let mut row = 2;
    for row_data in data {
        let region = row_data.get(0);
        if region != Some(&"East") && region != Some(&"North") {
            worksheet1.hide_row(row)?;
        }
        worksheet1.write_row((row, 1), row_data.iter())?;
        row += 1;
    }

    let mut filters = Filters::new();
    filters.or(Filter::eq("East")).or(Filter::eq("North"));
    worksheet1.filter_column("A1", &filters);

    workbook.save_as("examples/autofilter.xlsx")?;

    Ok(())
}