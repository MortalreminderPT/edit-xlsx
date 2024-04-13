use crate::utils::col_helper::{to_col, to_col_name, to_loc, to_ref};

pub(crate) trait Location {
    ///
    /// Convert the coordinates into a form like A1
    ///
    fn to_ref(&self) -> String;
    ///
    /// Convert the coordinates into a form like (1, 1)
    ///
    fn to_location(&self) -> (u32, u32);
    fn to_col(&self) -> u32;
    fn to_row(&self) -> u32;
}

pub(crate) trait LocationRange {
    ///
    /// Convert the coordinate range into a form such as A1:B6
    ///
    fn to_range_ref(&self) -> String;
    ///
    /// Convert the coordinate range into a form such as (1, 1, 2, 6)
    ///
    fn to_range(&self) -> (u32, u32, u32, u32);
    ///
    /// Convert the row coordinates to a form such as 1:6
    ///
    fn to_row_range_ref(&self) -> String;
    ///
    /// Convert the row coordinates to a form such as (1,6)
    ///
    fn to_row_range(&self) -> (u32, u32);
    ///
    /// Convert the col coordinates to a form such as A:B
    ///
    fn to_col_range_ref(&self) -> String;
    ///
    /// Convert the row coordinates to a form such as (1,2)
    ///
    fn to_col_range(&self) -> (u32, u32);
    ///
    /// Get the start point of range
    ///
    fn start_ref(&self) -> String;
    ///
    /// Get the end point of range
    ///
    fn end_ref(&self) -> String;
}

impl Location for &str {
    fn to_ref(&self) -> String {
        self.to_string()
    }

    fn to_location(&self) -> (u32, u32) {
        to_loc(self)
    }

    fn to_col(&self) -> u32 {
        to_col(self)
    }

    fn to_row(&self) -> u32 {
        let row = self.chars().filter(|&c| { c >= '0' && c <= '9' }).collect::<String>();
        row.parse().unwrap()
    }
}

impl Location for &String {
    fn to_ref(&self) -> String {
        self.to_string()
    }

    fn to_location(&self) -> (u32, u32) {
        to_loc(self)
    }

    fn to_col(&self) -> u32 {
        to_col(self)
    }

    fn to_row(&self) -> u32 {
        let row = self.chars().filter(|&c| { c >= '0' && c <= '9' }).collect::<String>();
        row.parse().unwrap()
    }
}

impl Location for (u32, u32) {
    fn to_ref(&self) -> String {
        to_ref(self.0, self.1)
    }

    fn to_location(&self) -> (u32, u32) {
        *self
    }

    fn to_col(&self) -> u32 {
        self.1
    }

    fn to_row(&self) -> u32 {
        self.0
    }
}

impl LocationRange for &str {
    fn to_range_ref(&self) -> String {
        self.to_string()
    }

    fn to_range(&self) -> (u32, u32, u32, u32) {
        let locs = self.split_once(':').unwrap();
        let start = locs.0.to_location();
        let end = locs.1.to_location();
        (start.0, start.1, end.0, end.1)
    }

    fn to_row_range_ref(&self) -> String {
        let locs = self.split_once(':').unwrap();
        let start_row = locs.0.chars().filter(|&c| c >= '0' && c <= '9').collect::<String>();
        let end_row = locs.1.chars().filter(|&c| c >= '0' && c <= '9').collect::<String>();
        format!("{}:{}", start_row, end_row)
    }

    fn to_row_range(&self) -> (u32, u32) {
        let locs = self.split_once(':').unwrap();
        let start_row: u32 = locs.0.chars().filter(|&c| c >= '0' && c <= '9').collect::<String>().parse().unwrap();
        let end_row: u32 = locs.1.chars().filter(|&c| c >= '0' && c <= '9').collect::<String>().parse().unwrap();
        (start_row, end_row)
    }

    fn to_col_range_ref(&self) -> String {
        let locs = self.split_once(':').unwrap();
        let start_col = locs.0.chars().filter(|&c| c >= 'A' && c <= 'Z').collect::<String>();
        let end_col = locs.1.chars().filter(|&c| c >= 'A' && c <= 'Z').collect::<String>();
        format!("{}:{}", start_col, end_col)
    }

    fn to_col_range(&self) -> (u32, u32) {
        let locs = self.split_once(':').unwrap();
        let start_col = locs.0.chars().filter(|&c| c >= 'A' && c <= 'Z').collect::<String>();
        let end_col = locs.1.chars().filter(|&c| c >= 'A' && c <= 'Z').collect::<String>();
        (to_col(&start_col), to_col(&end_col))
    }

    fn start_ref(&self) -> String {
        let locs = self.split_once(':').unwrap();
        locs.0.to_string()
    }

    fn end_ref(&self) -> String {
        let locs = self.split_once(':').unwrap();
        locs.1.to_string()
    }
}

impl LocationRange for &String {
    fn to_range_ref(&self) -> String {
        self.to_string()
    }

    fn to_range(&self) -> (u32, u32, u32, u32) {
        let locs = self.split_once(':').unwrap();
        let start = locs.0.to_location();
        let end = locs.1.to_location();
        (start.0, start.1, end.0, end.1)
    }

    fn to_row_range_ref(&self) -> String {
        let locs = self.split_once(':').unwrap();
        let start_row = locs.0.chars().filter(|&c| c >= '0' && c <= '9').collect::<String>();
        let end_row = locs.1.chars().filter(|&c| c >= '0' && c <= '9').collect::<String>();
        format!("{}:{}", start_row, end_row)
    }

    fn to_row_range(&self) -> (u32, u32) {
        let locs = self.split_once(':').unwrap();
        let start_row: u32 = locs.0.chars().filter(|&c| c >= '0' && c <= '9').collect::<String>().parse().unwrap();
        let end_row: u32 = locs.1.chars().filter(|&c| c >= '0' && c <= '9').collect::<String>().parse().unwrap();
        (start_row, end_row)
    }

    fn to_col_range_ref(&self) -> String {
        let locs = self.split_once(':').unwrap();
        let start_col = locs.0.chars().filter(|&c| c >= 'A' && c <= 'Z').collect::<String>();
        let end_col = locs.1.chars().filter(|&c| c >= 'A' && c <= 'Z').collect::<String>();
        format!("{}:{}", start_col, end_col)
    }

    fn to_col_range(&self) -> (u32, u32) {
        let locs = self.split_once(':').unwrap();
        let start_col = locs.0.chars().filter(|&c| c >= 'A' && c <= 'Z').collect::<String>();
        let end_col = locs.1.chars().filter(|&c| c >= 'A' && c <= 'Z').collect::<String>();
        (to_col(&start_col), to_col(&end_col))
    }

    fn start_ref(&self) -> String {
        let locs = self.split_once(':').unwrap();
        locs.0.to_string()
    }

    fn end_ref(&self) -> String {
        let locs = self.split_once(':').unwrap();
        locs.1.to_string()
    }
}

impl LocationRange for (u32, u32, u32, u32) {
    fn to_range_ref(&self) -> String {
        format!("{}{}:{}{}", to_col_name(self.1), self.0, to_col_name(self.3), self.2)
    }

    fn to_range(&self) -> (u32, u32, u32, u32) {
        *self
    }

    fn to_row_range_ref(&self) -> String {
        format!("{}:{}", self.0, self.2)
    }

    fn to_row_range(&self) -> (u32, u32) {
        (self.0, self.2)
    }

    fn to_col_range_ref(&self) -> String {
        format!("{}:{}", to_col_name(self.1), to_col_name(self.3))
    }

    fn to_col_range(&self) -> (u32, u32) {
        (self.1, self.3)
    }

    fn start_ref(&self) -> String {
        format!("{}{}", to_col_name(self.1), self.0)
    }

    fn end_ref(&self) -> String {
        format!("{}{}", to_col_name(self.3), self.2)
    }
}