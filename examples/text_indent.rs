use edit_xlsx::{WorkSheetCol, Format, FormatAlignType, Workbook, WorkbookResult, Write};

fn main() -> WorkbookResult<()> {
    // Create a new workbook
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet_mut(1)?;

    let indent1 = Format::default().set_indent(1);
    let indent2 = Format::default().set_indent(2);

    worksheet.set_columns_width("A:A", 40.0)?;

    worksheet.write_with_format("A1", "This text is indented 1 level", &indent1)?;
    worksheet.write_with_format("A2", "This text is indented 2 levels", &indent2)?;

    // Note: Alignment is not applied correctly when changing the reading order, this bug will be fixed in the future!
    // let indent = Format::default().set_reading_order(2).set_align(FormatAlignType::Right).set_indent(2);
    // worksheet.right_to_left();

    workbook.save_as("examples/text_indent.xlsx")?;
    Ok(())
}