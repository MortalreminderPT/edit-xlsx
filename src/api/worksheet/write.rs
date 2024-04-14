use std::slice::Iter;
use crate::api::cell::Cell;
use crate::api::cell::formula::Formula;
use crate::api::cell::values::{CellDisplay, CellValue};
use crate::api::cell::location::{Location, LocationRange};
use crate::api::worksheet::format::_Format;
use crate::api::worksheet::hyperlink::_Hyperlink;
use crate::Format;
use crate::api::worksheet::WorkSheet;
use crate::result::WorkSheetResult;
use crate::xml::extension::{AddExtension, ExtensionType};

pub trait Write: _Write {
    fn write_cell<L: Location, T: Clone + CellDisplay + CellValue>(&mut self, loc: L, cell: &Cell<T>) -> WorkSheetResult<()> {
        self.write_by_api_cell(&loc, &cell)?;
        Ok(())
    }

    fn write<L: Location, T: CellDisplay + CellValue>(&mut self, loc: L, data: T) -> WorkSheetResult<()> {
        self.write_display_all(&loc, &data, None)
    }
    fn write_string<L: Location>(&mut self, loc: L, data: String) -> WorkSheetResult<()> {
        let mut cell = Cell::default();
        cell.text = Some(data);
        self.write_by_api_cell(&loc, &cell)
    }
    fn write_number<L: Location>(&mut self, loc: L, data: i32) -> WorkSheetResult<()> {
        let mut cell = Cell::default();
        cell.text = Some(data);
        self.write_by_api_cell(&loc, &cell)
    }
    fn write_double<L: Location>(&mut self, loc: L, data: f64) -> WorkSheetResult<()> {
        let mut cell = Cell::default();
        cell.text = Some(data);
        self.write_by_api_cell(&loc, &cell)
    }
    fn write_boolean<L: Location>(&mut self, loc: L, data: bool) -> WorkSheetResult<()> {
        let mut cell = Cell::default();
        cell.text = Some(data);
        self.write_by_api_cell(&loc, &cell)
    }
    fn write_row<L: Location, T: CellDisplay + CellValue>(&mut self, loc: L, data: &[T]) -> WorkSheetResult<()> {
        let (row, mut col) = loc.to_location();
        for data in data {
            self.write_display_all(&(row, col), data, None)?;
            col += 1;
        }
        Ok(())
    }

    fn write_row_cells<L: Location, T: CellDisplay + CellValue + Clone>(&mut self, loc: L, cells: &[Cell<T>]) -> WorkSheetResult<()> {
        let (row, mut col) = loc.to_location();
        for cell in cells {
            self.write_by_api_cell(&(row, col), cell)?;
            col += 1;
        }
        Ok(())
    }

    fn write_column<L: Location, T: CellDisplay + CellValue>(&mut self, loc: L, data: &[T]) -> WorkSheetResult<()> {
        let (mut row, col) = loc.to_location();
        for data in data {
            self.write_display_all(&(row, col), data, None)?;
            row += 1;
        }
        Ok(())
    }

    fn write_column_cells<L: Location, T: CellDisplay + CellValue + Clone>(&mut self, loc: L, cells: &[Cell<T>]) -> WorkSheetResult<()> {
        let (mut row, col) = loc.to_location();
        for cell in cells {
            self.write_by_api_cell(&(row, col), cell)?;
            row += 1;
        }
        Ok(())
    }

    fn write_url<L: Location>(&mut self, loc: L, url: &str) -> WorkSheetResult<()> {
        let mut cell = Cell::default();
        cell.text = Some(url);
        cell.hyperlink = Some(url.to_string());
        self.write_by_api_cell(&loc, &cell)
    }

    fn write_url_text<L: Location, T: CellDisplay + CellValue>(&mut self, loc: L, url: &str, data: &str) -> WorkSheetResult<()> {
        let mut cell = Cell::default();
        cell.text = Some(data);
        cell.hyperlink = Some(url.to_string());
        self.write_by_api_cell(&loc, &cell)
    }

    fn merge_range<L: LocationRange, T: CellDisplay + CellValue>(&mut self, loc: L, data: T) -> WorkSheetResult<()> {
        self.merge_range_all(loc, data, None)
    }

    fn write_formula<L: Location>(&mut self, loc: L, data: &str) -> WorkSheetResult<()> {
        let mut cell: Cell<&str> = Cell::default();
        // cell.formula = Some(data.to_string());
        // cell.formula_type = Some("array".to_string());
        // cell.formula_ref = Some(loc.to_ref());
        cell.formula = Some(Formula::new_array_formula(data, &loc));
        // Some(FormulaType::Formula(loc.to_ref()).to_formula_ref());
        self.write_by_api_cell(&loc, &cell)
        // self.write_formula_all(&loc, data, FormulaType::Formula(loc.to_ref()), None)
    }

    fn write_old_formula<L: Location>(&mut self, loc: L, data: &str) -> WorkSheetResult<()> {
        let mut cell: Cell<&str> = Cell::default();
        cell.formula = Some(Formula::new_array_formula(data, &loc));
        // cell.formula_type = Some(FormulaType::OldFormula(loc.to_ref()));
        self.write_by_api_cell(&loc, &cell)
    }
    fn write_array_formula<L: Location>(&mut self, loc: L, data: &str) -> WorkSheetResult<()> {
        let mut cell: Cell<&str> = Cell::default();
        cell.formula = Some(Formula::new_array_formula(data, &loc));
        // cell.formula_type = Some(FormulaType::ArrayFormula(loc.to_ref()));
        self.write_by_api_cell(&loc, &cell)
    }
    fn write_dynamic_array_formula<L: Location>(&mut self, loc: L, data: &str) -> WorkSheetResult<()> {
        let mut cell: Cell<&str> = Cell::default();
        cell.formula = Some(Formula::new_array_formula(data, &loc));
        // cell.formula_type = Some(FormulaType::DynamicArrayFormula(loc.to_ref()));
        self.write_by_api_cell(&loc, &cell)
    }
    fn write_with_format<L: Location, T: Default + Clone + CellDisplay + CellValue>(&mut self, loc: L, data: T, format: &Format) -> WorkSheetResult<()> {
        let mut cell = Cell::default();
        cell.text = Some(data);
        cell.format = Some(format.clone());
        self.write_by_api_cell(&loc, &cell)
    }

    fn write_string_with_format<L: Location>(&mut self, loc: L, data: String, format: &Format) -> WorkSheetResult<()> {
        let mut cell = Cell::default();
        cell.text = Some(data);
        cell.format = Some(format.clone());
        self.write_by_api_cell(&loc, &cell)
    }

    fn write_number_with_format<L: Location>(&mut self, loc: L, data: i32, format: &Format) -> WorkSheetResult<()> {
        let mut cell = Cell::default();
        cell.text = Some(data);
        cell.format = Some(format.clone());
        self.write_by_api_cell(&loc, &cell)
    }
    fn write_double_with_format<L: Location>(&mut self, loc: L, data: f64, format: &Format) -> WorkSheetResult<()> {
        let mut cell = Cell::default();
        cell.text = Some(data);
        cell.format = Some(format.clone());
        self.write_by_api_cell(&loc, &cell)
    }
    fn write_boolean_with_format<L: Location>(&mut self, loc: L, data: bool, format: &Format) -> WorkSheetResult<()> {
        let mut cell = Cell::default();
        cell.text = Some(data);
        cell.format = Some(format.clone());
        self.write_by_api_cell(&loc, &cell)
    }
    fn write_row_with_format<L: Location, T: CellDisplay + CellValue>(&mut self, loc: L, data: Iter<'_, T>, format: &Format) -> WorkSheetResult<()> {
        let (row, mut col) = loc.to_location();
        for data in data {
            self.write_display_all(&(row, col), data, Some(format))?;
            col += 1;
        }
        Ok(())
    }
    fn write_column_with_format<L: Location, T: CellDisplay + CellValue>(&mut self, loc: L, data: Iter<'_, T>, format: &Format) -> WorkSheetResult<()> {
        let (mut row, col) = loc.to_location();
        for data in data {
            self.write_display_all(&(row, col), data, Some(format))?;
            row += 1;
        }
        Ok(())
    }
    fn write_url_with_format<L: Location>(&mut self, loc: L, url: &str, format: &Format) -> WorkSheetResult<()> {
        let mut cell = Cell::default();
        cell.text = Some(url);
        cell.hyperlink = Some(url.to_string());
        cell.format = Some(format.clone());
        self.write_by_api_cell(&loc, &cell)
    }
    fn write_url_text_with_format<L: Location>(&mut self, loc: L, url: &str, data: &str, format: &Format) -> WorkSheetResult<()> {
        let mut cell = Cell::default();
        cell.text = Some(data);
        cell.hyperlink = Some(url.to_string());
        cell.format = Some(format.clone());
        self.write_by_api_cell(&loc, &cell)
    }
    fn write_formula_with_format<L: Location>(&mut self, loc: L, data: &str, format: &Format) -> WorkSheetResult<()> {
        let mut cell: Cell<&str> = Cell::default();
        cell.formula = Some(Formula::new_array_formula(data, &loc));
        // cell.formula_type = Some(FormulaType::Formula(loc.to_ref()));
        cell.format = Some(format.clone());
        self.write_by_api_cell(&loc, &cell)
    }
    fn write_array_formula_with_format<L: Location>(&mut self, loc: L, data: &str, format: &Format) -> WorkSheetResult<()> {
        let mut cell: Cell<&str> = Cell::default();
        cell.formula = Some(Formula::new_array_formula(data, &loc));
        // cell.formula_type = Some(FormulaType::ArrayFormula(loc.to_ref()));
        cell.format = Some(format.clone());
        self.write_by_api_cell(&loc, &cell)
    }
    fn write_dynamic_array_formula_with_format<L: LocationRange>(&mut self, loc_range: L, data: &str, format: &Format) -> WorkSheetResult<()> {
        let loc = loc_range.to_range();
        let mut cell: Cell<&str> = Cell::default();
        cell.formula = Some(Formula::new_array_formula_by_range(data, &loc));
        cell.format = Some(format.clone());
        // cell.formula_type = Some(FormulaType::DynamicArrayFormula(loc_range.to_range_ref()));
        self.write_by_api_cell(&(loc.0, loc.1), &cell)
    }
    fn merge_range_with_format<L: LocationRange, T: CellDisplay + CellValue>(&mut self, loc: L, data: T, format:&Format) -> WorkSheetResult<()> {
        self.merge_range_all(loc, data, Some(format))
    }
}

trait _Write: _Format + _Hyperlink {
    fn write_by_api_cell<L: Location, T: CellDisplay + CellValue + Clone>(&mut self, loc: &L, cell: &Cell<T>) -> WorkSheetResult<()>;
    fn write_display_all<L: Location, T: CellDisplay + CellValue>(&mut self, loc: &L, data: &T, format: Option<&Format>) -> WorkSheetResult<()>;
    // fn write_formula_all<L: Location>(&mut self, loc: &L, formula: &str, formula_type: FormulaType, format: Option<&Format>) -> WorkSheetResult<()>;
    // fn write_hyperlink<L: Location>(&mut self, loc: &L, url: &str, data: &str, format: Option<&Format>) -> WorkSheetResult<()>;
    fn merge_range_all<L: LocationRange, T: CellDisplay + CellValue>(&mut self, loc: L, data: T, format: Option<&Format>) -> WorkSheetResult<()>;
}

impl _Write for WorkSheet {
    fn write_by_api_cell<L: Location, T: CellDisplay + CellValue + Clone>(&mut self, loc: &L, cell: &Cell<T>) -> WorkSheetResult<()> {
        let mut cell = cell.clone();
        if let Some(_) = &cell.formula {
            self.worksheet.xmlns_attrs.add_xr();
            self.worksheet.xmlns_attrs.add_xr_2();
            self.worksheet.xmlns_attrs.add_xr_3();
        }
        if let Some(format) = &cell.format {
            let style = self.add_format(format);
            cell.style = Some(style);
        }
        if let Some(url) = &cell.hyperlink {
            let url_r_id = self.worksheet_rel.add_hyperlink(url);
            self.worksheet.add_hyperlink(loc, url_r_id);
        }
        if let Some(_) = &cell.formula {
            // let formula_type = cell.formula_type
            //     .get_or_insert(
            //         // FormulaType::Formula(
            //         //     cell.formula_ref.clone().unwrap_or(loc.to_ref())
            //         // )
            //     );
            self.metadata.borrow_mut().add_extension(ExtensionType::XdaDynamicArrayProperties);
            self.workbook_rel.borrow_mut().get_or_add_metadata();
            self.content_types.borrow_mut().add_metadata();
            // if FormulaType::OldFormula(loc.to_ref()) != *formula_type {
            //     self.metadata.borrow_mut().add_extension(ExtensionType::XdaDynamicArrayProperties);
            // }
        }
        self.worksheet.sheet_data.write_by_api_cell(loc, &cell)?;
        Ok(())
    }

    fn write_display_all<L: Location, T: CellDisplay + CellValue>(&mut self, loc: &L, data: &T, format: Option<&Format>) -> WorkSheetResult<()> {
        let mut style = self.worksheet.get_default_style(loc);
        if let Some(format) = format {
            style = Some(self.add_format(format));
        }
        let worksheet = &mut self.worksheet;
        let sheet_data = &mut worksheet.sheet_data;
        sheet_data.write_display(loc, data, style)?;
        Ok(())
    }

    // fn write_formula_all<L: Location>(&mut self, loc: &L, formula: &str, formula_type: FormulaType, format: Option<&Format>) -> WorkSheetResult<()> {
    //     let mut style = None;
    //     if let Some(format) = format {
    //         style = Some(self.add_format(format));
    //     }
    //     let worksheet = &mut self.worksheet;
    //     let sheet_data = &mut worksheet.sheet_data;
    //     // self.workbook_rel.borrow_mut().get_or_add_metadata();
    //     self.workbook_rel.borrow_mut().get_or_add_metadata();
    //     self.content_types.borrow_mut().add_metadata();
    //     if FormulaType::OldFormula(loc.to_ref()) != formula_type {
    //         self.metadata.borrow_mut().add_extension(ExtensionType::XdaDynamicArrayProperties);
    //     }
    //     sheet_data.write_formula(loc, formula, formula_type, style)?;
    //     Ok(())
    // }

    // fn write_hyperlink<L: Location>(&mut self, loc: &L, url: &str, data: &str, format: Option<&Format>) -> WorkSheetResult<()> {
    //     self.write_display_all(loc, &data, format)?;
    //     let r_id = self.worksheet_rel.add_hyperlink(url);
    //     self.worksheet.add_hyperlink(loc, r_id);
    //     Ok(())
    // }

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