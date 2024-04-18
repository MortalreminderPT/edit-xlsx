#[cfg(test)]
mod tests {
    use edit_xlsx::{Format, FormatBorder, FormatBorderElement, FormatBorderType, FormatColor, Workbook, WorkbookResult, WorkSheetCol, WorkSheetRow};

    #[test]
    fn test_new() -> WorkbookResult<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.get_worksheet_mut(1)?;
        let mut format = Format::default();
        format.border = FormatBorder::default();
        format.border.left = FormatBorderElement::from_color(&FormatColor::RGB(255, 100, 100));
        format.border.right = FormatBorderElement::from_color(&FormatColor::RGB(255, 100, 100));
        worksheet.set_columns_width_with_format("B:C", 8.0, &format)?;
        format.border.top = FormatBorderElement::from_border_type(&FormatBorderType::Double);
        format.border.bottom = FormatBorderElement::from_border_type(&FormatBorderType::Double);
        worksheet.set_row_height_with_format(3, 15.0, &format)?;
        worksheet.set_row_height_with_format(4, 15.0, &format)?;
        worksheet.set_row_height_with_format(5, 15.0, &format)?;
        workbook.save_as("tests/output/border_test_new.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from() -> WorkbookResult<()> {
        let mut workbook = Workbook::from_path("tests/xlsx/shift-schedule.xlsx")?;
        let worksheet = workbook.get_worksheet_mut_by_name("Schedule")?;
        let mut format = Format::default();
        format.border = FormatBorder::default();
        format.border.left = FormatBorderElement::from_color(&FormatColor::RGB(255, 100, 100));
        format.border.right = FormatBorderElement::from_color(&FormatColor::RGB(255, 100, 100));
        worksheet.set_columns_width_with_format("A:U", 10.0, &format)?;
        format.border.top = FormatBorderElement::from_border_type(&FormatBorderType::Double);
        format.border.bottom = FormatBorderElement::from_border_type(&FormatBorderType::Double);
        worksheet.set_row_height_with_format(27, 15.0, &format)?;
        worksheet.set_row_height_with_format(4, 15.0, &format)?;
        worksheet.set_row_height_with_format(5, 15.0, &format)?;
        workbook.save_as("tests/output/border_test_from.xlsx")?;
        Ok(())
    }
}