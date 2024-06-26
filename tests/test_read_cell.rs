#[cfg(test)]
mod tests {
    use edit_xlsx::{Read, Workbook, WorkbookResult, WorkSheetCol, WorkSheetRow, Write};

    #[test]
    fn test_from() -> WorkbookResult<()> {
        // Read an existed workbook
        let reading_book = Workbook::from_path("./tests/xlsx/accounting.xlsx")?;
        let reading_sheet = reading_book.get_worksheet(1)?;
        // Create a new workbook to write
        let mut writing_book = Workbook::new();
        let writing_sheet = writing_book.get_worksheet_mut(1)?;

        // Synchronous column width and format
        let columns_map = reading_sheet.get_columns_with_format("A:XFD")?;
        writing_sheet.set_default_column(reading_sheet.get_default_column().unwrap_or(12.25));
        columns_map.iter().for_each(|(col_range, (column, format))| {
            if let Some(format) = format {
                // if col format exists, write it to writing_sheet
                writing_sheet.set_columns_with_format(col_range, column, format).unwrap()
            } else {
                writing_sheet.set_columns(col_range, column).unwrap()
            }
        });

        // Synchronous row height and format
        writing_sheet.set_default_row(reading_sheet.get_default_row());
        for row_number in 1..=reading_sheet.max_row() {
            let (row, format) = reading_sheet.get_row_with_format(row_number)?;
            if let Some(format) = format {
                // if col format exists, write it to writing_sheet
                writing_sheet.set_row_with_format(row_number, &row, &format)?;
            } else {
                writing_sheet.set_row(row_number, &row)?;
            }
        }

        // Read then write text and format
        for row in 1..=reading_sheet.max_row() {
            for col in 1..=reading_sheet.max_column() {
                if let Ok(cell) = reading_sheet.read_cell((row, col)) {
                    println!("{:?}", cell);
                    writing_sheet.write_cell((row, col), &cell)?;
                }
            }
        }
        writing_book.save_as("./tests/output/read_cell_test_from.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_calender() -> WorkbookResult<()> {
        // Read an existed workbook
        let mut calendar_workbook = Workbook::from_path("./tests/xlsx/yearly-calendar.xlsx")?;
        let worksheet = calendar_workbook.get_worksheet_mut_by_name("Calendar")?;
        worksheet.write_row("C3", &["hello", "world", "hello", "rust"])?;
        let cell = worksheet.read_cell("C31")?;
        calendar_workbook.save_as("./tests/output/read_cell_test_from_yearly_calendar.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_monthly_calender() -> WorkbookResult<()> {
        // Read an existed workbook
        let mut calendar_workbook = Workbook::from_path("./tests/xlsx/monthly-calendar.xlsx")?;
        let worksheet = calendar_workbook.get_worksheet_mut_by_name("Jan")?;
        let cell = worksheet.read_cell("C31")?;
        worksheet.write_column("A14", &["hello", "world", "hello", "rust"])?;
        println!("{:?}", cell);
        calendar_workbook.save_as("./tests/output/read_cell_test_from_monthly_calendar.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_week_calender() -> WorkbookResult<()> {
        // Read an existed workbook
        let mut calendar_workbook = Workbook::from_path("./tests/xlsx/week-calendar.xlsx")?;
        let worksheet = calendar_workbook.get_worksheet_mut_by_name("Calendar")?;
        worksheet.write_column("B6", &["hello", "world", "hello", "rust"])?;
        let cell = worksheet.read_cell("C31")?;
        println!("{:?}", cell);
        calendar_workbook.save_as("./tests/output/read_cell_test_from_week_calender.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_weekly_schedule() -> WorkbookResult<()> {
        // Read an existed workbook
        let mut schedule_workbook = Workbook::from_path("./tests/xlsx/weekly-schedule.xlsx")?;
        let worksheet = schedule_workbook.get_worksheet_mut_by_name("Week with hours")?;
        worksheet.write_column("D8", &["hello", "world", "hello", "rust"])?;
        let cell = worksheet.read_cell("C31")?;
        println!("{:?}", cell);
        schedule_workbook.save_as("./tests/output/read_cell_test_from_weekly_schedule.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_work_schedule() -> WorkbookResult<()> {
        // Read an existed workbook
        let mut schedule_workbook = Workbook::from_path("./tests/xlsx/work-schedule.xlsx")?;
        let worksheet = schedule_workbook.get_worksheet_mut_by_name("Week 1-2")?;
        worksheet.write_column("C6", &["hello", "world", "hello", "rust"])?;
        let cell = worksheet.read_cell("C31")?;
        println!("{:?}", cell);
        schedule_workbook.save_as("./tests/output/read_cell_test_from_work_schedule.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_shift_schedule() -> WorkbookResult<()> {
        // Read an existed workbook
        let mut schedule_workbook = Workbook::from_path("./tests/xlsx/shift-schedule.xlsx")?;
        let worksheet = schedule_workbook.get_worksheet_mut_by_name("Schedule")?;
        worksheet.write_row("B9", &["hello", "world", "hello", "rust"])?;
        let cell = worksheet.read_cell("C28")?;
        println!("{:?}", cell);
        #[cfg(feature = "ansi_term_support")]
        {
            use ansi_term::{ANSIStrings};
            println!("{}", ANSIStrings(&cell.ansi_strings()));
        }
        let cell = worksheet.read_cell("A8")?;
        println!("{:?}", cell);
        #[cfg(feature = "ansi_term_support")]
        {
            use ansi_term::{ANSIStrings};
            println!("{}", ANSIStrings(&cell.ansi_strings()));
        }
        let cell = worksheet.read_cell("A1")?;
        println!("{:?}", cell);
        #[cfg(feature = "ansi_term_support")]
        {
            use ansi_term::{ANSIStrings};
            println!("{}", ANSIStrings(&cell.ansi_strings()));
        }
        schedule_workbook.save_as("./tests/output/read_cell_test_from_shift_schedule.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_budget() -> WorkbookResult<()> {
        // Read an existed workbook
        let mut schedule_workbook = Workbook::from_path("./tests/xlsx/home-budget.xlsx")?;
        let worksheet = schedule_workbook.get_worksheet_mut_by_name("Budget")?;
        worksheet.write_column("B14", &(100..=107).into_iter().collect::<Vec<i32>>())?;
        let cell = worksheet.read_cell("C31")?;
        println!("{:?}", cell);
        schedule_workbook.save_as("./tests/output/read_cell_test_from_home_budget.xlsx")?;

        let mut schedule_workbook = Workbook::from_path("./tests/xlsx/personal-budget.xlsx")?;
        let worksheet = schedule_workbook.get_worksheet_mut_by_name("Budget")?;
        worksheet.write_column("B12", &(100..=105).into_iter().collect::<Vec<i32>>())?;
        let cell = worksheet.read_cell("C31")?;
        worksheet.insert_image("B12:E19", &"./examples/pics/capybara.bmp")?;
        println!("{:?}", cell);
        schedule_workbook.save_as("./tests/output/read_cell_test_from_personal_budget.xlsx")?;

        let mut schedule_workbook = Workbook::from_path("./tests/xlsx/wedding-budget.xlsx")?;
        let worksheet = schedule_workbook.get_worksheet_mut_by_name("Breakdown")?;
        worksheet.write_column("B12", &(100..=105).into_iter().collect::<Vec<i32>>())?;
        let cell = worksheet.read_cell("C31")?;
        worksheet.insert_image("B12:E19", &"./examples/pics/capybara.bmp")?;
        println!("{:?}", cell);
        schedule_workbook.save_as("./tests/output/read_cell_test_from_wedding_budget.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_register() -> WorkbookResult<()> {
        // Read an existed workbook
        let mut schedule_workbook = Workbook::from_path("./tests/xlsx/checkbook-register.xlsx")?;
        let worksheet = schedule_workbook.get_worksheet_mut_by_name("Register")?;
        worksheet.write_row("B9", &["hello", "world", "hello", "rust"])?;
        let cell = worksheet.read_cell("C31")?;
        for i in 6..16 {
            worksheet.set_row_level(i, 1)?;
            worksheet.collapse_row(i)?;
        }
        println!("{:?}", cell);
        schedule_workbook.save_as("./tests/output/read_cell_test_from_checkbook_register.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_meeting() -> WorkbookResult<()> {
        // Read an existed workbook
        let mut schedule_workbook = Workbook::from_path("./tests/xlsx/world-meeting-planner.xlsx")?;
        // let worksheet = schedule_workbook.get_worksheet_mut_by_name("Register")?;
        schedule_workbook.save_as("./tests/output/read_cell_test_from_world_meeting_planner.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_chart() -> WorkbookResult<()> {
        // Read an existed workbook
        let mut org_workbook = Workbook::from_path("./tests/xlsx/company-organization-chart.xlsx")?;
        let worksheet = org_workbook.get_worksheet_mut_by_name("WithImages")?;
        worksheet.insert_image("A16:E25", &"./examples/pics/capybara.bmp")?;
        worksheet.insert_image("A30:E40", &"./examples/pics/ferris.png")?;
        org_workbook.save_as("./tests/output/read_cell_test_from_company_organization_chart.xlsx")?;
        Ok(())
    }

    #[test]
    fn test_from_calculator() -> WorkbookResult<()> {
        // Read an existed workbook
        let mut calculator_workbook = Workbook::from_path("./tests/xlsx/paycheck-calculator.xlsx")?;
        let worksheet = calculator_workbook.get_worksheet_mut_by_name("NEW W-4")?;
        worksheet.write("C5", 1234567)?;
        // worksheet.write("C5", 1234567)?;
        worksheet.write_old_formula("F50", "=IF(C11=\"Yes\",0,IF(C10=\"Married Joint\",12900,8600))")?;

        calculator_workbook.save_as("./tests/output/read_cell_test_from_paycheck_calculator.xlsx")?;
        Ok(())
    }
}