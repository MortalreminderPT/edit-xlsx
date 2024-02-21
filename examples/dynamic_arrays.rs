use edit_xlsx::{Col, Format, FormatColor, Workbook, WorkbookResult, WorkSheet, WorkSheetResult, Write};

fn main() -> WorkbookResult<()> {
    // Create a new workbook called simple.xls and add some worksheets.
    let mut workbook = Workbook::new();
    let header1 = Format::default()
        .set_color(FormatColor::RGB("00ffffff"))
        .set_background_color(FormatColor::RGB("0074AC4C"));
    let header2 = Format::default()
        .set_color(FormatColor::RGB("00ffffff"))
        .set_background_color(FormatColor::RGB("00528FD3"));

    //
    // Example of using the FILTER() function.
    //
    let worksheet1 = workbook.add_worksheet_by_name("Filter")?;
    worksheet1.write_formula("F2", "_xlfn.FILTER(A1:D17,C1:C17=K2)")?;

    // Write the data the function will work on.
    worksheet1.write_with_format("K1", "Product", &header2)?;
    worksheet1.write("K2", "Apple")?;
    worksheet1.write_with_format("F1", "Region", &header2)?;
    worksheet1.write_with_format("G1", "Sales Rep", &header2)?;
    worksheet1.write_with_format("H1", "Product", &header2)?;
    worksheet1.write_with_format("I1", "Units", &header2)?;

    write_worksheet_data(worksheet1, &header1)?;
    worksheet1.set_column_pixels("E:E", 10.0)?;
    worksheet1.set_column_pixels("J:J", 10.0)?;

    //
    // Example of using the UNIQUE() function.
    //
    let worksheet2 = workbook.add_worksheet_by_name("Unique")?;
    worksheet2.write_formula("F2", "_xlfn.UNIQUE(B2:B17)")?;

    // A more complex example combining SORT and UNIQUE.
    worksheet2.write_formula("H2", "_xlfn.SORT(_xlfn.UNIQUE(B2:B17))")?;

    // Write the data the function will work on.
    worksheet2.write_with_format("F1", "Sales Rep", &header2)?;
    worksheet2.write_with_format("H1", "Sales Rep", &header2)?;

    write_worksheet_data(worksheet2, &header1)?;
    worksheet2.set_column_pixels("E:E", 10.0)?;
    worksheet2.set_column_pixels("G:G", 10.0)?;

    //
    // Example of using the SORT() function.
    //
    let worksheet3 = workbook.add_worksheet_by_name("Sort")?;
    worksheet3.write_formula("F2", "_xlfn.SORT(B2:B17)")?;

    // A more complex example combining SORT and FILTER.
    worksheet3.write_formula("H2", "_xlfn.SORT(_xlfn.FILTER(C2: D17, D2: D17 > 5000, \"\"), 2, 1)")?;

    // Write the data the function will work on.
    worksheet3.write_with_format("F1", "Sales Rep", &header2)?;
    worksheet3.write_with_format("H1", "Product", &header2)?;
    worksheet3.write_with_format("I1", "Units", &header2)?;

    write_worksheet_data(worksheet3, &header1)?;
    worksheet3.set_column_pixels("E:E", 10.0)?;
    worksheet3.set_column_pixels("G:G", 10.0)?;


    //
    // Example of using the SORTBY() function.
    //
    let worksheet4 = workbook.add_worksheet_by_name("Sortby")?;
    worksheet4.write_formula("D2", "_xlfn.SORTBY(A2:B9,B2:B9)")?;

    // Write the data the function will work on.
    worksheet4.write_with_format("A1", "Name", &header1)?;
    worksheet4.write_with_format("B1", "Age", &header1)?;

    worksheet4.write("A2", "Tom")?;
    worksheet4.write("A3", "Fred")?;
    worksheet4.write("A4", "Amy")?;
    worksheet4.write("A5", "Sal")?;
    worksheet4.write("A6", "Fritz")?;
    worksheet4.write("A7", "Srivan")?;
    worksheet4.write("A8", "Xi")?;
    worksheet4.write("A9", "Hector")?;

    worksheet4.write("B2", 52)?;
    worksheet4.write("B3", 65)?;
    worksheet4.write("B4", 22)?;
    worksheet4.write("B5", 73)?;
    worksheet4.write("B6", 19)?;
    worksheet4.write("B7", 39)?;
    worksheet4.write("B8", 19)?;
    worksheet4.write("B9", 66)?;

    worksheet4.write_with_format("D1", "Name", &header2)?;
    worksheet4.write_with_format("E1", "Age", &header2)?;

    worksheet4.set_column_pixels("C:C", 10.0)?;

    //
    // Example of using the XLOOKUP() function.
    //
    let worksheet5 = workbook.add_worksheet_by_name("Xlookup")?;
    worksheet5.write_formula("F1", "_xlfn.XLOOKUP(E1,A2:A9,C2:C9)")?;

    // Write the data the function will work on.
    worksheet5.write_with_format("A1", "Country", &header1)?;
    worksheet5.write_with_format("B1", "Abr", &header1)?;
    worksheet5.write_with_format("C1", "Prefix", &header1)?;

    worksheet5.write("A2", "China")?;
    worksheet5.write("A3", "India")?;
    worksheet5.write("A4", "United States")?;
    worksheet5.write("A5", "Indonesia")?;
    worksheet5.write("A6", "Brazil")?;
    worksheet5.write("A7", "Pakistan")?;
    worksheet5.write("A8", "Nigeria")?;
    worksheet5.write("A9", "Bangladesh")?;

    worksheet5.write("B2", "CN")?;
    worksheet5.write("B3", "IN")?;
    worksheet5.write("B4", "US")?;
    worksheet5.write("B5", "ID")?;
    worksheet5.write("B6", "BR")?;
    worksheet5.write("B7", "PK")?;
    worksheet5.write("B8", "NG")?;
    worksheet5.write("B9", "BD")?;

    worksheet5.write("C2", 86)?;
    worksheet5.write("C3", 91)?;
    worksheet5.write("C4", 1)?;
    worksheet5.write("C5", 62)?;
    worksheet5.write("C6", 55)?;
    worksheet5.write("C7", 92)?;
    worksheet5.write("C8", 234)?;
    worksheet5.write("C9", 880)?;

    worksheet5.write_with_format("E1", "Brazil", &header2)?;

    worksheet5.set_column_pixels("A:A", 100.0)?;
    worksheet5.set_column_pixels("D:D", 10.0)?;

    //
    // Example of using the XMATCH() function.
    //
    let worksheet6 = workbook.add_worksheet_by_name("Xmatch")?;
    worksheet6.write_formula("D2", "_xlfn.XMATCH(C2,A2:A6)")?;

    // Write the data the function will work on.
    worksheet6.write_with_format("A1", "Product", &header1)?;

    worksheet6.write("A2", "Apple")?;
    worksheet6.write("A3", "Grape")?;
    worksheet6.write("A4", "Pear")?;
    worksheet6.write("A5", "Banana")?;
    worksheet6.write("A6", "Cherry")?;

    worksheet6.write_with_format("C1", "Product", &header2)?;
    worksheet6.write_with_format("D1", "Position", &header2)?;
    worksheet6.write("C2", "Grape")?;

    worksheet6.set_column_pixels("B:B", 10.0)?;

    //
    // Example of using the RANDARRAY() function.
    //
    let worksheet7 = workbook.add_worksheet_by_name("Randarray")?;
    worksheet7.write_dynamic_array_formula("A1", "_xlfn.RANDARRAY(5,3,1,100, TRUE)")?;

    //
    // Example of using the SEQUENCE() function.
    //
    let worksheet8 = workbook.add_worksheet_by_name("Sequence")?;
    worksheet8.write_dynamic_array_formula("A1", "_xlfn.SEQUENCE(4,5)")?;

    //
    // Example of using the Spill range operator.
    //
    let worksheet9 = workbook.add_worksheet_by_name("Spill ranges")?;
    worksheet9.write_dynamic_array_formula("H2", "_xlfn.ANCHORARRAY(F2)")?;

    worksheet9.write_dynamic_array_formula("J2", "_xlfn.COUNTA(_xlfn.ANCHORARRAY(F2))")?;

    // Write the data the to work on.
    worksheet9.write_dynamic_array_formula("F2", "_xlfn.UNIQUE(B2:B17)")?;
    worksheet9.write_with_format("F1", "Unique", &header2)?;
    worksheet9.write_with_format("H1", "Spill", &header2)?;
    worksheet9.write_with_format("J1", "Spill", &header2)?;

    write_worksheet_data(worksheet9, &header1)?;
    worksheet9.set_column_pixels("E:E", 10.0)?;
    worksheet9.set_column_pixels("G:G", 10.0)?;
    worksheet9.set_column_pixels("I:I", 10.0)?;

    //
    // Example of using dynamic ranges with older Excel functions.
    //
    let worksheet10 = workbook.add_worksheet_by_name("Older functions")?;
    worksheet10.write_dynamic_array_formula("B1", "=LEN(A1:A3)")?;

    // Write the data the to work on.
    worksheet10.write("A1", "Foo")?;
    worksheet10.write("A2", "Food")?;
    worksheet10.write("A3", "Frood")?;


    workbook.save_as("examples/dynamic_arrays.xlsx")?;
    Ok(())
}

fn write_worksheet_data(worksheet: &mut WorkSheet, header: &Format) -> WorkSheetResult<()> {
    worksheet.write_with_format("A1", "Region", &header)?;
    worksheet.write_with_format("B1", "Sales Rep", &header)?;
    worksheet.write_with_format("C1", "Product", &header)?;
    worksheet.write_with_format("D1", "Units", &header)?;
    let data = [
        ["East", "Tom", "Apple"],
        ["West", "Fred", "Grape"],
        ["North", "Amy", "Pear"],
        ["South", "Sal", "Banana"],
        ["East", "Fritz", "Apple"],
        ["West", "Sravan", "Grape"],
        ["North", "Xi", "Pear"],
        ["South", "Hector", "Banana"],
        ["East", "Tom", "Banana"],
        ["West", "Fred", "Pear"],
        ["North", "Amy", "Grape"],
        ["South", "Sal", "Apple"],
        ["East", "Fritz", "Banana"],
        ["West", "Sravan", "Pear"],
        ["North", "Xi", "Grape"],
        ["South", "Hector", "Apple"],
    ];
    let mut row_num = 2;
    for data in data {
        worksheet.write_row((row_num, 1), data.iter())?;
        row_num += 1;
    }
    let units = [6380, 5619, 4565, 5323, 4394, 7195, 5231, 2427, 4213, 3239, 6520, 1310, 6274, 4894, 7580, 9814];
    worksheet.write_column("D2", units.iter())?;
    Ok(())
}