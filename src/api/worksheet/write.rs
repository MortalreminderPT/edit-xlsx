use std::ops::DerefMut;
use std::slice::Iter;
use crate::api::cell::formula::FormulaType;
use crate::api::cell::values::{CellDisplay, CellValue};
use crate::api::cell::location::{Location, LocationRange};
use crate::api::worksheet::format::_Format;
use crate::api::worksheet::hyperlink::_Hyperlink;
use crate::Format;
use crate::api::worksheet::WorkSheet;
use crate::result::WorkSheetResult;
use crate::xml::extension::{AddExtension, ExtensionType};

pub trait Write: _Write {
    fn write<L: Location, T: CellDisplay + CellValue>(&mut self, loc: L, data: T) -> WorkSheetResult<()> {
        self.write_display_all(&loc, &data, None)
    }
    fn write_string<L: Location>(&mut self, loc: L, data: String) -> WorkSheetResult<()> { self.write_display_all(&loc, &data, None) }
    fn write_number<L: Location>(&mut self, loc: L, data: i32) -> WorkSheetResult<()> { self.write_display_all(&loc, &data, None) }
    fn write_double<L: Location>(&mut self, loc: L, data: f64) -> WorkSheetResult<()> { self.write_display_all(&loc, &data, None) }
    fn write_boolean<L: Location>(&mut self, loc: L, data: bool) -> WorkSheetResult<()> { self.write_display_all(&loc, &data, None) }
    fn write_row<L: Location, T: CellDisplay + CellValue>(&mut self, loc: L, data: Iter<'_, T>) -> WorkSheetResult<()> {
        let (row, mut col) = loc.to_location();
        for data in data {
            self.write_display_all(&(row, col), data, None)?;
            col += 1;
        }
        Ok(())
    }
    fn write_column<L: Location, T: CellDisplay + CellValue>(&mut self, loc: L, data: Iter<'_, T>) -> WorkSheetResult<()> {
        let (mut row, col) = loc.to_location();
        for data in data {
            self.write_display_all(&(row, col), data, None)?;
            row += 1;
        }
        Ok(())
    }
    fn write_url<L: Location>(&mut self, loc: L, url: &str) -> WorkSheetResult<()> {
        self.write_hyperlink(&loc, url, url, None)
    }
    fn write_url_text<L: Location>(&mut self, loc: L, url: &str, data: &str) -> WorkSheetResult<()> {
        self.write_hyperlink(&loc, url, data, None)
    }
    fn merge_range<L: LocationRange, T: CellDisplay + CellValue>(&mut self, loc: L, data: T) -> WorkSheetResult<()> {
        self.merge_range_all(loc, data, None)
    }
    fn write_formula<L: Location>(&mut self, loc: L, data: &str) -> WorkSheetResult<()> { self.write_formula_all(&loc, data, FormulaType::Formula(loc.to_ref()), None) }
    fn write_array_formula<L: Location>(&mut self, loc: L, data: &str) -> WorkSheetResult<()> {
        // let loc = &loc_range.to_locations();
        self.write_formula_all(&loc, data, FormulaType::ArrayFormula(loc.to_ref()), None)
    }
    fn write_dynamic_array_formula<L: Location>(&mut self, loc: L, data: &str) -> WorkSheetResult<()> {
        self.write_formula_all(&loc, data, FormulaType::DynamicArrayFormula(loc.to_ref()), None)
    }

    fn write_with_format<L: Location, T: CellDisplay + CellValue>(&mut self, loc: L, data: T, format: &Format) -> WorkSheetResult<()> { self.write_display_all(&loc, &data, Some(format)) }
    fn write_string_with_format<L: Location>(&mut self, loc: L, data: String, format: &Format) -> WorkSheetResult<()> { self.write_display_all(&loc, &data, Some(format)) }
    fn write_number_with_format<L: Location>(&mut self, loc: L, data: i32, format: &Format) -> WorkSheetResult<()> { self.write_display_all(&loc, &data, Some(format)) }
    fn write_double_with_format<L: Location>(&mut self, loc: L, data: f64, format: &Format) -> WorkSheetResult<()> { self.write_display_all(&loc, &data, Some(format)) }
    fn write_boolean_with_format<L: Location>(&mut self, loc: L, data: bool, format: &Format) -> WorkSheetResult<()> { self.write_display_all(&loc, &data, Some(format)) }
    fn write_row_with_format<L: Location, T: CellDisplay + CellValue>(&mut self, loc: L, mut data: Vec<T>, format: &Format) -> WorkSheetResult<()> {
        let (row, col) = loc.to_location();
        let mut col = col + data.len() as u32 - 1;
        while !data.is_empty() {
            self.write_display_all(&(row, col), &data.pop().unwrap(), Some(format))?;
            col -= 1;
        }
        Ok(())
    }
    fn write_column_with_format<L: Location, T: CellDisplay + CellValue>(&mut self, loc: L, mut data: Vec<T>, format: &Format) -> WorkSheetResult<()> {
        let (row, col) = loc.to_location();
        let mut row = row + data.len() as u32 - 1;
        while !data.is_empty() {
            self.write_display_all(&(row, col), &data.pop().unwrap(), Some(format))?;
            row -= 1;
        }
        Ok(())
    }
    fn write_url_with_format<L: Location>(&mut self, loc: L, url: &str, format: &Format) -> WorkSheetResult<()> {
        self.write_hyperlink(&loc, url, url, Some(format))
    }
    fn write_url_data_with_format<L: Location>(&mut self, loc: L, url: &str, data: &str, format: &Format) -> WorkSheetResult<()> {
        self.write_hyperlink(&loc, url, data, Some(format))
    }
    fn write_formula_with_format<L: Location>(&mut self, loc: L, data: &str, format: &Format) -> WorkSheetResult<()> {
        self.write_formula_all(&loc, data, FormulaType::Formula(loc.to_ref()), Some(format))
    }
    fn write_array_formula_with_format<L: Location>(&mut self, loc: L, data: &str, format: &Format) -> WorkSheetResult<()> {
        self.write_formula_all(&loc.to_location(), data, FormulaType::ArrayFormula(loc.to_ref()), Some(format))
    }
    fn write_dynamic_array_formula_with_format<L: LocationRange>(&mut self, loc_range: L, data: &str, format: &Format) -> WorkSheetResult<()> {
        let loc = loc_range.to_range();
        self.write_formula_all(&(loc.0, loc.1), data, FormulaType::DynamicArrayFormula(loc_range.to_range_ref()), Some(format))
    }
    
    fn merge_range_with_format<L: LocationRange, T: CellDisplay + CellValue>(&mut self, loc: L, data: T, format:&Format) -> WorkSheetResult<()> {
        self.merge_range_all(loc, data, Some(format))
    }
}

pub(crate) trait _Write: _Format + _Hyperlink {
    fn write_display_all<L: Location, T: CellDisplay + CellValue>(&mut self, loc: &L, data: &T, format: Option<&Format>) -> WorkSheetResult<()>;
    fn write_formula_all<L: Location>(&mut self, loc: &L, formula: &str, formula_type: FormulaType, format: Option<&Format>) -> WorkSheetResult<()>;
    fn write_hyperlink<L: Location>(&mut self, loc: &L, url: &str, data: &str, format: Option<&Format>) -> WorkSheetResult<()>;
    fn merge_range_all<L: LocationRange, T: CellDisplay + CellValue>(&mut self, loc: L, data: T, format: Option<&Format>) -> WorkSheetResult<()>;
}

impl _Write for WorkSheet {
    fn write_display_all<L: Location, T: CellDisplay + CellValue>(&mut self, loc: &L, data: &T, format: Option<&Format>) -> WorkSheetResult<()> {
        let mut style = None;
        if let Some(format) = format {
            style = Some(self.add_format(format));
        }
        let worksheet = &mut self.worksheet;
        let sheet_data = &mut worksheet.sheet_data;
        sheet_data.write_display(loc, data, style)?;
        Ok(())
    }

    fn write_formula_all<L: Location>(&mut self, loc: &L, formula: &str, formula_type: FormulaType, format: Option<&Format>) -> WorkSheetResult<()> {
        let mut style = None;
        if let Some(format) = format {
            style = Some(self.add_format(format));
        }
        let worksheet = &mut self.worksheet;
        let sheet_data = &mut worksheet.sheet_data;
        // self.workbook_rel.borrow_mut().get_or_add_metadata();
        self.workbook_rel.borrow_mut().get_or_add_metadata();
        self.content_types.borrow_mut().add_metadata();
        self.metadata.borrow_mut().add_extension(ExtensionType::XdaDynamicArrayProperties);
        sheet_data.write_formula(loc, formula, formula_type, style)?;
        Ok(())
    }

    fn write_hyperlink<L: Location>(&mut self, loc: &L, url: &str, data: &str, format: Option<&Format>) -> WorkSheetResult<()> {
        self.write_display_all(loc, &data, format)?;
        let r_id = self.worksheet_rel.add_hyperlink(url);
        self.worksheet.add_hyperlink(loc, r_id);
        Ok(())
    }

    fn merge_range_all<L: LocationRange, T: CellDisplay + CellValue>(&mut self, loc: L, data: T, format: Option<&Format>) -> WorkSheetResult<()> {
        let (first_row, first_col, last_row, last_col) = loc.to_range();
        let worksheet = &mut self.worksheet;
        worksheet.add_merge_cell(first_row, first_col, last_row, last_col);
        for row in first_row..=last_row {
            for col in first_col..=last_col {
                self.write_display_all(&(row, col), &data, format)?;
            }
        }
        Ok(())
    }
}