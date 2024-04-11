use edit_xlsx::{Col, Format, FormatAlignType, FormatBorderType, FormatColor, Row, Workbook, WorkbookResult, Write};

fn main() -> WorkbookResult<()> {
    let header_format = Format::default()
        .set_bold()
        .set_align(FormatAlignType::Center)
        .set_align(FormatAlignType::VerticalCenter)
        .set_border(FormatBorderType::Medium)
        .set_background_color(FormatColor::RGB("00D7E4BC".to_string()));
    let center_format = Format::default().set_align(FormatAlignType::Center);

    // Create a new workbook
    let mut workbook = Workbook::new();

    //
    // Example 1. Freeze pane on the top row.
    //
    let worksheet1 = workbook.add_worksheet_by_name("Panes 1")?;
    worksheet1.freeze_panes("A2")?;
    // Other sheet formatting.
    worksheet1.set_column("A:I", 16.0)?;
    worksheet1.set_row(0, 20.0)?;
    worksheet1.set_selection("C3:C3")?;

    // Some text to demonstrate scrolling.
    for col in 1..=9 {
        worksheet1.write_with_format((1, col), "Scroll down", &header_format)?;
    }

    for row in 2..100 {
        for col in 1..=9 {
            worksheet1.write_with_format((row, col), row, &center_format)?;
        }
    }

    //
    // Example 2. Freeze pane on the left column.
    //
    let worksheet2 = workbook.add_worksheet_by_name("Panes 2")?;
    worksheet2.freeze_panes("B1")?;

    // Other sheet formatting.
    worksheet2.set_column("A:A", 16.0)?;
    worksheet2.set_selection("C3:C3")?;

    // Some text to demonstrate scrolling.
    for row in 1..=50 {
        worksheet2.write_with_format((row, 1), "Scroll right", &header_format)?;
        for col in 2..=26 {
            worksheet2.write_with_format((row, col), col, &center_format)?;
        }
    }

    //
    // Example 3. Freeze pane on the top row and left column.
    //
    let worksheet3 = workbook.add_worksheet_by_name("Panes 3")?;
    worksheet3.freeze_panes((2, 2))?;

    // Other sheet formatting.
    worksheet3.set_column("A:Z", 16.0)?;
    worksheet3.set_row(1, 20.0)?;
    worksheet3.set_selection("C3:C3")?;
    worksheet3.write_with_format((1, 1), "", &header_format)?;

    // Some text to demonstrate scrolling.
    for col in 2..=26 {
        worksheet3.write_with_format((1, col), "Scroll down", &header_format)?;
    }

    for row in 2..=50 {
        worksheet3.write_with_format((row, 1), "Scroll right", &header_format)?;
        for col in 2..=26 {
            worksheet3.write_with_format((row, col), col, &center_format)?;
        }
    }

    //
    // Example 4. Split pane on the top row and left column.
    //
    // The divisions must be specified in terms of row and column dimensions.
    //
    let worksheet4 = workbook.add_worksheet_by_name("Panes 4")?;
    // Set the default row height is 17 and set the column width is 13
    worksheet4.set_column("A:Z", 13.0)?;
    worksheet4.set_default_row(17.0);
    worksheet4.split_panes(2.0 * 13.0, 2.0 * 17.0)?;
    worksheet4.set_selection("C3:C3")?;

    // Some text to demonstrate scrolling.
    for col in 1..=26 {
        worksheet4.write_with_format((1, col), "Scroll", &center_format)?;
    }
    for row in 1..=50 {
        worksheet4.write_with_format((row, 1), "Scroll", &center_format)?;
        for col in 1..=26 {
            worksheet4.write_with_format((row, col), col, &center_format)?;
        }
    }

    workbook.save_as("examples/panes.xlsx")?;
    Ok(())
}