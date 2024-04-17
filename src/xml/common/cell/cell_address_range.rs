use crate::utils::col_helper;

#[derive(Default, Debug, Copy, Clone)]
struct CellAddressRange {
    first_row: u32,
    first_col: u32,
    last_row: u32,
    last_col: u32,
}

impl CellAddressRange {
    fn new(first_row: u32, first_col: u32, last_row: u32, last_col: u32) -> CellAddressRange {
        CellAddressRange {
            first_row,
            first_col,
            last_row,
            last_col,
        }
    }

    fn from_str(refer: &str) -> CellAddressRange {
        let mut cell_address_range = CellAddressRange::default();
        let cell_addresses: Vec<&str> = refer.split(":").collect();
        match cell_addresses.len() {
            1 => {
                (cell_address_range.first_row, cell_address_range.first_col) =
                    col_helper::to_loc(cell_addresses[0]);
                (cell_address_range.last_row, cell_address_range.last_col) =
                    (cell_address_range.first_row, cell_address_range.first_col)
            },
            2 => {
                (cell_address_range.first_row, cell_address_range.first_col) =
                    col_helper::to_loc(cell_addresses[0]);
                (cell_address_range.last_row, cell_address_range.last_col) =
                    col_helper::to_loc(cell_addresses[1]);
            },
            _ => {}
        }
        cell_address_range
    }

    fn to_string(&self) -> String {
        todo!()
        // return the string like A1:B2 format
    }

    fn to_cell_address_string(&self) -> String {
        todo!()
        // just return the start address string like A1 format
    }
}