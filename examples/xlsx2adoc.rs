use std::io::BufWriter;
use std::io::Error;
use std::io::Write;
use std::path::Path;

use std::fs::File;

use edit_xlsx::Read;
use edit_xlsx::WorkSheet;
use edit_xlsx::WorkSheetCol;
use edit_xlsx::{FormatColor, Workbook, WorkbookResult};
use std::ops::RangeInclusive;

use crate::helper::error_text;

use crate::tabheader::tab_header;

#[derive(Default)]
pub struct Xlsx2AdocTestResults {
    // Todo
    v1: u8,
    pub v2: u8,
}

pub(crate) fn to_col(col: &str) -> u32 {
    let mut col = col.as_bytes();
    let mut num = 0;
    while !col.is_empty() {
        if col[0] > 64 && col[0] < 91 {
            num *= 26;
            num += (col[0] - 64) as u32;
        }
        col = &col[1..];
    }
    num
}

fn decode_col_range(column_name: &str) -> RangeInclusive<u32> {
    let mut nn = column_name.split(':');
    let nl = nn.next();
    let nl = nl.unwrap();
    let cl = to_col(nl);
    let nh = nn.next();
    let nh = nh.unwrap();
    let ch = to_col(nh);
    cl - 1..=ch - 1
}

fn find_col_width(sheet: &WorkSheet) -> Result<Vec<f64>, Error> {
    let mut widths = Vec::<f64>::new();
    let default_col_width = sheet.get_default_column();

    for _ in 0..sheet.max_column() {
        widths.push(default_col_width);
    }

    let formatted_col_result = sheet.get_columns_with_format((1, 1, 1, 16384));
    let formatted_col = match formatted_col_result {
        Ok(f) => f,
        Err(e) => return Err(error_text(&format!("{:?}", e))),
    };

    for w in formatted_col.iter() {
        let column_name = w.0;
        let a = w.1;
        let columns_specs = a.0;
        let column_width = columns_specs.width;
        let column_width_string = match column_width {
            Some(width) => {
                let col_range = decode_col_range(column_name);
                for c in col_range {
                    widths[c as usize] = width;
                }
                format!("{}", width)
            }
            None => "-".to_owned(),
        };
        let format = &a.1;
        let x3 = format.clone();
        let comumn_format_fill = match x3 {
            Some(x) => {
                let x = x.get_background();
                format!("{:?}", x)
            }
            None => "-".to_owned(),
        };
    }

    Ok(widths)
}

pub fn xlsx_convert(
    in_file_name: &Path,
    out_file_name: &Path,
) -> Result<Xlsx2AdocTestResults, Error> {
    let workbook = Workbook::from_path(in_file_name);
    let mut workbook = workbook.unwrap();
    workbook.finish();

    let reading_sheet = workbook.get_worksheet(1);
    let sheet = reading_sheet.unwrap();
    let default_row_hight = sheet.get_default_row();

    println!(
        "Rows {} -> ? ( {} )",
        sheet.max_row(),
        //        formatted_row.len(),
        default_row_hight
    );

    let mut hights = Vec::<f64>::new();
    for _ in 0..sheet.max_row() {
        hights.push(default_row_hight);
    }

    let widths = find_col_width(sheet)?;
    let bounds = "|===\r";
    let line = tab_header(&widths);

    let mut output_file = File::create(out_file_name)?; // overwrites existing file
    let mut writer = BufWriter::new(&mut output_file);

    writer.write(line.as_bytes())?;
    writer.write(bounds.as_bytes())?;

    /*     // todo test
       writer.write("|1|2|3\r".as_bytes())?;
       writer.write("|4|5|6\r".as_bytes())?;

       writer.write(bounds.as_bytes())?;
    */
    for row in 0..sheet.max_row() {
        println!("Row {} ({})", row, hights[row as usize]);
        for col in 0..sheet.max_column() {
            if col < sheet.max_column() {
                writer.write("|".as_bytes())?;
            }

            let cell_content = sheet.read_cell((row + 1, col + 1)).unwrap_or_default();
            let format = cell_content.format;
            let mut text_color = FormatColor::Default;
            let mut bg_color = FormatColor::Default;
            let mut bg_bg_color = FormatColor::Default;
            if format.is_some() {
                let format = format.unwrap();
                text_color = format.get_color().clone();
                let ff = format.get_background().clone();
                bg_color = ff.fg_color.clone();
                bg_bg_color = ff.bg_color.clone();
            }

            let cell_format_string = format!(
                "Text-Color = {:?}        bg = {:?}        bg_bg = {:?}",
                text_color, bg_color, bg_bg_color
            );
            let cell_text = cell_content.text;
            let text = match cell_text {
                Some(t) => t,
                None => "-".to_owned(),
            };

            println!(
                "{} ({}) -> {}     Format: {}",
                col,
                widths[(col) as usize],
                text,
                cell_format_string
            );
            writer.write(text.as_bytes())?;
        }

        writer.write("\r".as_bytes())?;
    }
    writer.write(bounds.as_bytes())?;

    let xlsx_2_adoc_test_results = Xlsx2AdocTestResults { v1: 0, v2: 0 };
    Ok(xlsx_2_adoc_test_results)
}
