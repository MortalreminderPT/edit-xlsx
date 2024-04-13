use std::fs;
use edit_xlsx::{Col, Filter, Filters, Format, Row, Workbook, WorkbookResult, Write};

fn main() -> WorkbookResult<()> {
    // Prepare autofilter data
    let text = fs::read_to_string("examples/autofilter_data.txt").unwrap();
    let mut text = text.split("\n");
    let headers: Vec<&str> = text.next().unwrap().split_whitespace().collect();
    let mut data: Vec<Vec<&str>> = vec![];
    for text in text { data.push(text.split_whitespace().collect()) }

    // Create a new workbook
    let mut workbook = Workbook::new();
    // Add some worksheets
    workbook.add_worksheet()?;
    workbook.add_worksheet()?;
    workbook.add_worksheet()?;
    workbook.add_worksheet()?;
    workbook.add_worksheet()?;
    workbook.add_worksheet()?;

    // Set up several sheets with the same data.
    for worksheet in workbook.worksheets_mut() {
        // Make the columns wider.
        worksheet.set_columns_width("A:D", 12.0)?;
        // // Make the header row larger.
        worksheet.set_row_with_format(1, 20.0, &Format::default().set_bold())?;
        // Make the headers bold.
        worksheet.write_row("A1", &headers)?;
    }

    //
    // Example 1. Autofilter without conditions.
    //

    let worksheet1 = workbook.get_worksheet_mut(1)?;
    // Set the autofilter.
    worksheet1.autofilter("A1:D51");
    let mut row = 2;
    for row_data in &data {
        let mut col = 1;
        for data in row_data {
            if let Ok(num) = data.parse::<i32>() {
                worksheet1.write((row, col), num)?;
            } else {
                worksheet1.write((row, col), *data)?;
            }
            col += 1;
        }
        // Move on to the next worksheet row.
        row += 1;
    }


    //
    // Example 2. Autofilter with a filter condition in the first column.
    //
    let worksheet2 = workbook.get_worksheet_mut(2)?;
    // Set the autofilter.
    worksheet2.autofilter("A1:D51");
    // Add filter criteria.
    let mut filters = Filters::new();
    filters.and(Filter::eq("East"));
    worksheet2.filter_column("A", &filters);
    // Hide the rows that don't match the filter criteria.
    let mut row = 2;
    for row_data in &data {
        let mut col = 1;
        let data = row_data.get(0);
        // Check for rows that match the filter.
        if data != Some(&"East") {
            // We need to hide rows that don't match the filter.
            worksheet2.hide_row(row)?;
        }
        for data in row_data {
            if let Ok(num) = data.parse::<i32>() {
                worksheet2.write((row, col), num)?;
            } else {
                worksheet2.write((row, col), *data)?;
            }
            col += 1;
        }
        // Move on to the next worksheet row.
        row += 1;
    }

    //
    // Example 3. Autofilter with a filter condition in the first column.
    //
    let worksheet3 = workbook.get_worksheet_mut(3)?;
    // Set the autofilter.
    worksheet3.autofilter("A1:D51");
    // Add filter criteria.
    let mut filters = Filters::new();
    filters.and(Filter::eq("East")).or(Filter::eq("South"));
    worksheet3.filter_column("A", &filters);
    // Hide the rows that don't match the filter criteria.
    let mut row = 2;
    for row_data in &data {
        let mut col = 1;
        let data = row_data.get(0);
        // Check for rows that match the filter.
        if data != Some(&"East") && data != Some(&"South") {
            // We need to hide rows that don't match the filter.
            worksheet3.hide_row(row)?;
        }
        for data in row_data {
            if let Ok(num) = data.parse::<i32>() {
                worksheet3.write((row, col), num)?;
            } else {
                worksheet3.write((row, col), *data)?;
            }
            col += 1;
        }
        // Move on to the next worksheet row.
        row += 1;
    }


    //
    // Example 4. Autofilter with filter conditions in two columns.
    //
    let worksheet4 = workbook.get_worksheet_mut(4)?;
    // Set the autofilter.
    worksheet4.autofilter("A1:D51");
    // Add filter criteria.
    let mut filters_a = Filters::new();
    filters_a.and(Filter::eq("East"));
    worksheet4.filter_column("A", &filters_a);
    let mut filters_c = Filters::new();
    filters_c.and(Filter::gt("3000")).and(Filter::lt("8000"));
    worksheet4.filter_column("C", &filters_c);
    // Hide the rows that don't match the filter criteria.
    let mut row = 2;
    for row_data in &data {
        let mut col = 1;
        let data = row_data.get(0);
        // Check for rows that match the filter.
        if data != Some(&"East") {
            // We need to hide rows that don't match the filter.
            worksheet4.hide_row(row)?;
        }
        for data in row_data {
            if let Ok(num) = data.parse::<i32>() {
                if num <= 3000 || num >= 8000 {
                    worksheet4.hide_row(row)?;
                }
                worksheet4.write((row, col), num)?;
            } else {
                worksheet4.write((row, col), *data)?;
            }
            col += 1;
        }
        // Move on to the next worksheet row.
        row += 1;
    }


    //
    // Example 5. Autofilter with a filter list condition in one of the columns.
    //
    let worksheet5 = workbook.get_worksheet_mut(5)?;
    // Set the autofilter.
    worksheet5.autofilter("A1:D51");
    // Add filter criteria.
    let filters_list = Filters::eq(vec!["East", "North", "South"]);
    worksheet5.filter_column("A", &filters_list);
    // Hide the rows that don't match the filter criteria.
    let mut row = 2;
    for row_data in &data {
        let mut col = 1;
        let data = row_data.get(0);
        // Check for rows that match the filter.
        if data != Some(&"East") && data != Some(&"North") && data != Some(&"South") {
            // We need to hide rows that don't match the filter.
            worksheet5.hide_row(row)?;
        }
        for data in row_data {
            if let Ok(num) = data.parse::<i32>() {
                worksheet5.write((row, col), num)?;
            } else {
                worksheet5.write((row, col), *data)?;
            }
            col += 1;
        }
        // Move on to the next worksheet row.
        row += 1;
    }

    //
    // Example 6. Autofilter with filter for blanks.
    //
    let worksheet6 = workbook.get_worksheet_mut(6)?;
    // Set the autofilter.
    worksheet6.autofilter("A1:D51");
    // Add filter criteria.
    let filters = Filters::blank();
    worksheet6.filter_column("A", &filters);
    // Hide the rows that don't match the filter criteria.
    let mut row = 2;
    // Simulate a blank cell in the data.
    data[5][0] = "";

    for row_data in &data {
        let mut col = 1;
        let data = row_data.get(0);
        // Check for rows that match the filter.
        if data != Some(&"") {
            // We need to hide rows that don't match the filter.
            worksheet6.hide_row(row)?;
        }
        for data in row_data {
            if let Ok(num) = data.parse::<i32>() {
                worksheet6.write((row, col), num)?;
            } else {
                worksheet6.write((row, col), *data)?;
            }
            col += 1;
        }
        // Move on to the next worksheet row.
        row += 1;
    }

    //
    // Example 7. Autofilter with filter for non-blanks.
    //
    let worksheet7 = workbook.get_worksheet_mut(7)?;
    // Set the autofilter.
    worksheet7.autofilter("A1:D51");
    // Add filter criteria.
    let filters = Filters::not_blank();
    worksheet7.filter_column("A", &filters);
    // Hide the rows that don't match the filter criteria.
    let mut row = 2;
    // Simulate a blank cell in the data.

    for row_data in &data {
        let mut col = 1;
        let data = row_data.get(0);
        // Check for rows that match the filter.
        if data == Some(&"") {
            // We need to hide rows that don't match the filter.
            worksheet7.hide_row(row)?;
        }
        for data in row_data {
            if let Ok(num) = data.parse::<i32>() {
                worksheet7.write((row, col), num)?;
            } else {
                worksheet7.write((row, col), *data)?;
            }
            col += 1;
        }
        // Move on to the next worksheet row.
        row += 1;
    }


    workbook.save_as("examples/autofilter.xlsx")?;

    Ok(())
}