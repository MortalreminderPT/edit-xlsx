use crate::utils::col_helper::{to_col_name, to_loc, to_ref};

pub(crate) trait Location {
    fn to_ref(&self) -> String;
    fn to_location(&self) -> (u32, u32);
}

pub(crate) trait LocationRange {
    fn to_ref(&self) -> String;
    fn to_locations(&self) -> (u32, u32, u32, u32);
}

impl Location for &str {
    fn to_ref(&self) -> String {
        self.to_string()
    }

    fn to_location(&self) -> (u32, u32) {
        to_loc(self)
    }
}

impl Location for (u32, u32) {
    fn to_ref(&self) -> String {
        to_ref(self.0, self.1)
    }

    fn to_location(&self) -> (u32, u32) {
        *self
    }
}

impl LocationRange for &str {
    fn to_ref(&self) -> String {
        self.to_string()
    }

    fn to_locations(&self) -> (u32, u32, u32, u32) {
        let locs = self.split_once(':').unwrap();
        let start = locs.0.to_location();
        let end = locs.1.to_location();
        (start.0, start.1, end.0, end.1)
    }
}

impl LocationRange for (u32, u32, u32, u32) {
    fn to_ref(&self) -> String {
        format!("{}{}:{}{}", to_col_name(self.1), self.0, to_col_name(self.3), self.2)
    }

    fn to_locations(&self) -> (u32, u32, u32, u32) {
        *self
    }
}