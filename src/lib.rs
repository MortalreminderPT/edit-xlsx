mod workbook;
mod sheet;
mod row;
mod cell;
mod col;
mod file;
mod xml;
mod result;
mod utils;

pub use workbook::Workbook;

fn hello() {
    println!("hello");
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_hello() {
        hello();
    }
}
