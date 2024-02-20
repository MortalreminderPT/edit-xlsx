use std::collections::HashMap;
use edit_xlsx::{FormatColor, Workbook, WorkbookResult, Write};

fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::new();
    // let mut workbook = Workbook::from_path("new.xlsx")?;
    let worksheet = workbook.get_worksheet(1)?;
    // add some worksheets
    workbook.add_worksheet()?;
    workbook.add_worksheet_by_name("Foglio2")?;
    workbook.add_worksheet_by_name("Data")?;
    workbook.add_worksheet()?;
    for worksheet in workbook.worksheets() {
        // get worksheet id and name
        worksheet.write("A1", format!("text in sheet{}, {}", worksheet.id(), worksheet.get_name()))?;
    }
    // get worksheet by name
    let worksheet = workbook.get_worksheet_by_name("Data")?;
    // set background for sheet
    worksheet.set_background(&"./examples/pics/ferris.png");
    // change tab color of sheet
    let tab_color = FormatColor::RGB("00ff0000");
    worksheet.set_tab_color(&tab_color);
    // worksheet.set_zoom(200);
    worksheet.set_top_left_cell("A1");
    worksheet.set_selection("A1:B2");
    worksheet.set_selection("C1:D5");
    worksheet.set_selection("F5");
    worksheet.freeze_panes("A1");
    // worksheet.set_top_left_cell("AB128")
    worksheet.activate();
    // worksheet.hide_unused_rows(true);
    // worksheet.set_default_row(100.0);
    worksheet.write("B2", format!("text in sheet{}, {}", worksheet.id(), worksheet.get_name()))?;
    worksheet.write("C10", format!("text in sheet{}, {}", worksheet.id(), worksheet.get_name()))?;
    // worksheet.outline_settings(false, false, false, true);

    for i in (1..100) {
        for j in (1..100) {
            worksheet.write((i, j), format!("{i}{j}"))?;
        }
    }

    // worksheet.write_string("A1", "123546".to_string())?;
    // let mut map = HashMap::new();
    // map.insert("number_stored_as_text", "A1:XFD1048576");
    // map.insert("formula_differs", "A1:XFD1048576");
    // map.insert("eval_error", "A1:XFD1048576");
    // worksheet.ignore_errors(map);
    println!("{}", worksheet.max_column());
    println!("{}", worksheet.max_row());
    workbook.save_as("new_3.xlsx")?;
    Ok(())
}