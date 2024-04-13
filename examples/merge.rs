use edit_xlsx::{Format, Workbook, WorkbookResult, FormatColor, FormatAlignType, WorkSheetCol, Write, Row, FormatBorderType};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet_mut(1)?;

    // Increase the cell size of the merged cells to highlight the formatting.
    worksheet.set_columns_width("B:D", 18.0)?;
    worksheet.set_row(4, 40.0)?;
    worksheet.set_row(7, 30.0)?;
    worksheet.set_row(8, 30.0)?;

    // Create a format to use in the merged range.
    let merge_format = Format::default()
        .set_bold()
        .set_border(FormatBorderType::Double)
        .set_align(FormatAlignType::Center)
        .set_align(FormatAlignType::VerticalCenter)
        .set_background_color(FormatColor::RGB(255, 255, 0));

    // Merge 3 cells.
    worksheet.merge_range_with_format("B4:D4", "Merged Range", &merge_format)?;
    // Merge 3 cells over two rows.
    worksheet.merge_range_with_format("B7:D8", "Merged Range", &merge_format)?;

    workbook.save_as("examples/merge.xlsx")?;
    Ok(())
}